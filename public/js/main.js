import { setContactToFailedList } from "./modules/setContactToFailedList.js";

/** No Google Sheet — change this to your real Meet / demo URL after audio check */
const DEMO_MEETING_LINK = "https://meet.google.com/";
const DEMO_ROOM_COUNT = 1;

function debugLog(msg) {
  console.log(`[${new Date().toISOString()}] ${msg}`);
}

// ELEMENTS
const timerContainer = document.querySelector(".countdown");
const signInContainer = document.querySelector(".signIn");
const noEduguest = document.querySelector(".no-eduquest");
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
const spinners = document.querySelectorAll(".spinner");
const headerTitles = document.querySelectorAll(".heading__title");
const signInButton = document.querySelector(".signIn-google");
const issuesAudioLink = document.querySelector(".issues-audio-link");

// ETC VARIABLES
let roomNumber = DEMO_ROOM_COUNT;
let links = [DEMO_MEETING_LINK];
let googleName = "";
/** After Join succeeded — blocks any late UI from flipping back to the form */
let joinFlowStarted = false;

const joinButton = form.elements.joinButton;
joinButton.disabled = true;

// PARAMS
export function getParamValue(param) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(param);
}

function showSignInForm() {
  if (joinFlowStarted) {
    debugLog("showSignInForm skipped (joinFlowStarted) ");
    return;
  }

  debugLog("showSignInForm: showing name/email form");
  headerTitles.forEach((title) => {
    title.textContent = "Tech check";
  });

  localStorage.setItem("failureLink", DEMO_MEETING_LINK);

  timerContainer.style.display = "none";
  noEduguest.style.display = "none";
  signInContainer.style.display = "block";
  joinButton.disabled = false;
}

// Log native validation failures (no submit handler run if invalid)
form.addEventListener(
  "invalid",
  (e) => {
    const t = e.target;
    debugLog(
      `Native validation blocked submit: ${t.name || t.id} — ${t.validationMessage}`
    );
  },
  true
);

// "JOIN"-FORM SUBMIT
form.addEventListener("submit", onJoinFormSubmit);

async function onJoinFormSubmit(evt) {
  evt.preventDefault();

  debugLog("Join clicked: submit handler running");

  const name = form.elements.name.value;
  const processName = name.split(" ").length > 1 ? name : "";
  const email = form.elements.email.value;
  const loginCredential = email ? email : getUserACId();

  const room = await getRandomNumber(roomNumber);
  const meetingLink = links[room - 1];
  debugLog(`room=${room} meetingLink=${meetingLink ? "ok" : "MISSING"}`);
  if (!meetingLink) {
    debugLog("ABORT: no meeting link for room");
    console.error("No meeting link for assigned room.");
    return;
  }

  joinButton.disabled = true;

  fetch("/isEduquestActive")
    .then((response) => {
      debugLog(`/isEduquestActive status=${response.status}`);
      if (!response.ok) {
        throw new Error(`isEduquestActive HTTP ${response.status}`);
      }
      return response.json();
    })
    .then(async (res) => {
      const data = res.data;
      debugLog(`isEduquestActive data[0]=${data && data[0]} redirect=${data && data[1]}`);
      if (!data) {
        throw new Error("isEduquestActive: missing data");
      }
      if (data[0] === true) {
        const urlId = generateIdForURL();
        debugLog(`urlId=${urlId} (after generateIdForURL)`);

        const response = await fetch("/users");
        const users = await response.json();

        const user = users.find((user) => user.id === urlId);
        if (!user) {
          const postRes = await fetch("/users", {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              id: urlId,
              googleName: googleName || " ",
              name: processName || name.trim() || "Guest",
              loginCredential,
              meetingLink,
              mainRoomNumber: room,
            }),
          });
          if (!postRes.ok) {
            const errBody = await postRes.json().catch(() => ({}));
            debugLog(`POST /users FAILED: ${JSON.stringify(errBody)}`);
            console.error("Could not create user:", errBody);
            joinButton.disabled = false;
            return;
          }
          debugLog("POST /users OK");
        } else {
          debugLog("user already exists in memory");
        }

        try {
          if (typeof io !== "function") {
            throw new Error(
              "Socket.IO client not loaded. Use /socket.io/socket.io.js from this server."
            );
          }
          socketConnection(urlId);
          debugLog("socketConnection() OK");
        } catch (err) {
          debugLog(`socketConnection ERROR: ${err.message}`);
          console.error(err);
          joinButton.disabled = false;
          return;
        }

        try {
          setQRCodeElements();
          debugLog("QRCode init OK");
        } catch (err) {
          debugLog(`QRCode init skipped (non-fatal): ${err.message}`);
          console.error(err);
        }

        joinFlowStarted = true;
        techCheckContainer.style.display = "block";
        signInContainer.style.display = "none";
        joinButton.disabled = false;
        techCheckContainer.scrollIntoView({ behavior: "smooth", block: "start" });
        debugLog("SUCCESS: tech-check UI visible (Prepare Your Tech) — scrolled into view");
      } else {
        debugLog(`redirecting to baselink: ${data[1]}`);
        window.location.href = data[1];
      }
    })
    .catch((error) => {
      debugLog(`Join flow ERROR: ${error.message}`);
      console.log(error.message);
      joinButton.disabled = false;
    });
}

function setQRCodeElements() {
  const cameraQRcode = new QRCode(document.getElementById("cameraQRcode"), {
    text: generateURL("camera"),
  });

  const microphoneQRcode = new QRCode(
    document.getElementById("microphoneQRcode"),
    {
      text: generateURL("microphone"),
    }
  );
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
    const currentURL = window.location.href;
    window.location.href = `${currentURL}&techCheck=failed`;
  });

  socket.on("microphoneCheckPassed", () => {
    techMicroContainer.style.display = "none";
    techAudioContainer.style.display = "flex";
  });

  socket.on("microphoneCheckFailed", () => {
    const currentURL = window.location.href;
    window.location.href = `${currentURL}&techCheck=failed`;
  });
}

async function getRandomNumber(maxNumber) {
  const response = await fetch("/getNextRoomNumber", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      maxNumber,
    }),
  });
  const { roomNumber: nextRoom } = await response.json();
  return nextRoom;
}

function generateIdForURL() {
  const currentUrl = window.location.href;
  const idRegex = /\/(\d+)(?:\?.*)?/;
  const hasId = currentUrl.match(idRegex);

  if (!hasId) {
    const searchParams = new URLSearchParams(window.location.search);

    const pathname = window.location.pathname.endsWith("/")
      ? window.location.pathname.slice(0, -1)
      : window.location.pathname;

    const id = generateUniqueId();
    const newUrl = `${
      window.location.origin
    }${pathname}/${id}?${searchParams.toString()}&generatedId=true`;
    window.history.replaceState({}, "", newUrl);
    return id;
  }

  return hasId[1];
}

function generateUniqueId() {
  return (
    new Date().getTime().toString(36).slice(-6) +
    Math.random().toString(36).slice(2, 8)
  );
}

// TECH-CHECK FORM SUBMIT
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
  const { meetingLink } = await response.json();

  await fetch(`/users/${userACId}`, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      camera: true,
      microphone: true,
      audio: true,
      meetingLink: meetingLink,
    }),
  });

  audioForm.reset();
  window.location.href = meetingLink;
}

function isValidInputValue(input) {
  const validPattern = /^(21|twenty[ -]?one|2[ -]?1)$/;
  //twenty-one, twenty one, 21, twentyone, 2 1

  return validPattern.test(input);
}

function getUserACId() {
  const currentUrl = window.location.href;
  const match = currentUrl.match(/\/(\w+)(?:\?.*)?$/);
  if (match) {
    return match[1];
  }
  return null;
}

// GOOGLE SIGN IN
signInButton.addEventListener("click", handleGoogleSignIn);

function handleGoogleSignIn() {
  const currentURL = window.location.href;
  const parts = currentURL.split("/");
  const userId = parts[parts.length - 1];

  const redirectUrl = `/signin/google?userId=${userId}`;
  window.location.href = redirectUrl;
}

// URL PARAMS LOGIC
async function paramsInterfaceLogic() {
  const techCheck = getParamValue("techCheck");
  const signInSuccess = getParamValue("signInSuccess");
  const generatedId = getParamValue("generatedId");

  debugLog(
    `init: techCheck=${techCheck} signInSuccess=${signInSuccess} generatedId=${generatedId} path=${window.location.pathname}`
  );

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
    techCheckContainer.style.display = "block";
  }
}

paramsInterfaceLogic();

function handleTechCheckFailed() {
  failureLink.href = localStorage.getItem("failureLink") || DEMO_MEETING_LINK;
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

//SET CONTACT TO THE TECH CHECK LIST
failureLink.addEventListener("click", setContactToFailedList);
issuesAudioLink.addEventListener("click", setContactToFailedList);
