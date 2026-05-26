import { setContactToFailedList } from "./modules/setContactToFailedList.js";

function debugLog(msg) {
  console.log(`[${new Date().toISOString()}] ${msg}`);
}

const signInContainer = document.querySelector(".signIn");
const techCheckContainer = document.querySelector(".tech-check");
const techCameraContainer = document.querySelector(".tech-camera-container");
const techMicroContainer = document.querySelector(".tech-microphone-container");
const anotherDeviceText = document.querySelector("#qr-microphone-text");
const failureText = document.getElementById("failure-text");
const failureLink = document.querySelector(".tech-failure-link");
const techAudioContainer = document.querySelector(".tech-audio-container");
const audioForm = document.querySelector(".audio-form");
const audioButton = document.querySelector("#audio-check-btn");
const audioFailureTextEl = document.querySelector(".audio-failure-text");
const form = document.querySelector(".form");
const emailInput = document.querySelector("#emailAddress");
const headerTitles = document.querySelectorAll(".heading__title");
const signInButton = document.querySelector(".signIn-google");
const issuesAudioLink = document.querySelector(".issues-audio-link");

let appConfig = { meetingLink: "", failureLink: "" };
let googleName = "";
let joinFlowStarted = false;

const joinButton = form.elements.joinButton;
joinButton.disabled = true;

export function getParamValue(param) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(param);
}

async function loadAppConfig() {
  const response = await fetch("/api/config");
  if (!response.ok) {
    throw new Error(`Failed to load config: HTTP ${response.status}`);
  }
  appConfig = await response.json();
  localStorage.setItem("failureLink", appConfig.failureLink);
  if (issuesAudioLink) {
    issuesAudioLink.href = appConfig.failureLink;
  }
}

function showSignInForm() {
  if (joinFlowStarted) {
    return;
  }

  headerTitles.forEach((title) => {
    title.textContent = "Tech check";
  });

  signInContainer.style.display = "block";
  joinButton.disabled = false;
}

form.addEventListener("submit", onJoinFormSubmit);

async function onJoinFormSubmit(evt) {
  evt.preventDefault();

  const name = form.elements.name.value;
  const processName = name.split(" ").length > 1 ? name : "";
  const email = form.elements.email.value;

  joinButton.disabled = true;

  try {
    await loadAppConfig();

    if (!appConfig.meetingLink) {
      console.error("Meeting link is not configured. Ask admin to set it at /admin");
      joinButton.disabled = false;
      return;
    }

    const urlId = getOrCreateUserId();
    const loginCredential = email || urlId;

    const postRes = await fetch("/users", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        id: urlId,
        googleName: googleName || " ",
        name: processName || name.trim() || "Guest",
        loginCredential,
      }),
    });
    if (!postRes.ok) {
      const errBody = await postRes.json().catch(() => ({}));
      console.error("Could not create user:", errBody);
      joinButton.disabled = false;
      return;
    }

    markUrlAsJoined(urlId);

    if (typeof io !== "function") {
      throw new Error("Socket.IO client not loaded.");
    }
    socketConnection(urlId);
    setQRCodeElements();

    joinFlowStarted = true;
    techCheckContainer.style.display = "block";
    signInContainer.style.display = "none";
    joinButton.disabled = false;
    techCheckContainer.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    debugLog(`Join flow ERROR: ${error.message}`);
    console.error(error);
    joinButton.disabled = false;
  }
}

function setQRCodeElements() {
  new QRCode(document.getElementById("cameraQRcode"), {
    text: generateURL("camera"),
  });

  new QRCode(document.getElementById("microphoneQRcode"), {
    text: generateURL("microphone"),
  });
}

function generateURL(parameter) {
  const currentURL = window.location.href;
  return `${currentURL}&${parameter}=true`;
}

function socketConnection(urlId) {
  const socket = io();

  socket.on("connect", () => {
    socket.emit("join", urlId);
  });

  socket.on("cameraCheckPassed", () => {
    techCameraContainer.style.display = "none";
    techMicroContainer.style.display = "block";
    anotherDeviceText.style.display = "none";
  });

  socket.on("cameraCheckFailed", () => {
    window.location.href = `${window.location.href}&techCheck=failed`;
  });

  socket.on("microphoneCheckPassed", () => {
    techMicroContainer.style.display = "none";
    techAudioContainer.style.display = "flex";
  });

  socket.on("microphoneCheckFailed", () => {
    window.location.href = `${window.location.href}&techCheck=failed`;
  });
}

function getUserACId() {
  const segment = window.location.pathname.split("/").filter(Boolean)[0];
  if (!segment || segment === "admin") {
    return null;
  }
  return segment;
}

function getOrCreateUserId() {
  const existingId = getUserACId();
  if (existingId) {
    return existingId;
  }

  const id = generateUniqueId();
  const searchParams = new URLSearchParams(window.location.search);
  const newUrl = `${window.location.origin}/${id}?${searchParams.toString()}`;
  window.history.replaceState({}, "", newUrl);
  return id;
}

function markUrlAsJoined(urlId) {
  const searchParams = new URLSearchParams(window.location.search);
  searchParams.set("generatedId", "true");
  window.history.replaceState(
    {},
    "",
    `${window.location.origin}/${urlId}?${searchParams.toString()}`
  );
}

function generateUniqueId() {
  return (
    new Date().getTime().toString(36).slice(-6) +
    Math.random().toString(36).slice(2, 8)
  );
}

audioForm.addEventListener("submit", onAudioFormSubmit);

async function onAudioFormSubmit(evt) {
  evt.preventDefault();
  const inputValue = evt.target.elements.audioCheck.value.trim().toLowerCase();
  if (!isValidInputValue(inputValue)) {
    audioFailureTextEl.style.display = "block";
    audioForm.reset();
    return;
  }
  audioButton.disabled = true;

  const userACId = getUserACId();
  const response = await fetch(`/users/${userACId}`);
  if (!response.ok) {
    console.error("User not found. Please join again.");
    audioButton.disabled = false;
    return;
  }
  const user = await response.json();

  await fetch(`/users/${userACId}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      camera: true,
      microphone: true,
      audio: true,
    }),
  });

  audioForm.reset();
  window.location.href = user.meetingLink;
}

function isValidInputValue(input) {
  return /^(21|twenty[ -]?one|2[ -]?1)$/.test(input);
}

signInButton.addEventListener("click", handleGoogleSignIn);

function handleGoogleSignIn() {
  const userId = getUserACId();
  const query = userId ? `?userId=${encodeURIComponent(userId)}` : "";
  window.location.href = `/signin/google${query}`;
}

async function paramsInterfaceLogic() {
  const techCheck = getParamValue("techCheck");
  const signInSuccess = getParamValue("signInSuccess");
  const generatedId = getParamValue("generatedId");

  try {
    await loadAppConfig();
  } catch (error) {
    console.error(error);
  }

  if (signInSuccess && !techCheck && !generatedId) {
    handleSuccessSignIn();
  }
  if (!techCheck && !generatedId) {
    showSignInForm();
  }
  if (techCheck) {
    handleTechCheckFailed();
  }
  if (generatedId) {
    const urlId = getUserACId();
    if (urlId) {
      try {
        const res = await fetch(`/users/${urlId}`);
        if (!res.ok) {
          showSignInForm();
          return;
        }
        joinFlowStarted = true;
        techCheckContainer.style.display = "block";
        signInContainer.style.display = "none";
        socketConnection(urlId);
        setQRCodeElements();
      } catch (error) {
        console.error(error);
        showSignInForm();
      }
    }
  }
}

paramsInterfaceLogic();

function handleTechCheckFailed() {
  failureLink.href = appConfig.failureLink || localStorage.getItem("failureLink") || "#";
  signInButton.style.display = "none";
  signInContainer.style.display = "none";
  techCameraContainer.style.display = "none";
  techCheckContainer.style.display = "block";
  failureText.style.display = "block";
}

function handleSuccessSignIn() {
  googleName = getParamValue("googleName");
  signInButton.style.display = "none";

  if (getUserACId()) {
    emailInput.removeAttribute("required");
    emailInput.removeAttribute("pattern");
    emailInput.parentNode.classList.add("hidden");
  }
}

failureLink.addEventListener("click", setContactToFailedList);
issuesAudioLink.addEventListener("click", setContactToFailedList);
