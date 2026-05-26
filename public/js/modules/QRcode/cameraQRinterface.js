import { QRvideoInspector } from "./QRvideoInspector.js";
import { setContactToFailedList } from "../setContactToFailedList.js";

const cameraCheckButtonEl = document.getElementById("camera-check-btn");
const qrProcceedBtnEl = document.querySelector(".qr-button");
const qrFailureMessage = document.querySelector(".qr-failure-message");
const cameraFailureEl = document.getElementById("camera-failure-text");
const videoContainerEl = document.querySelector(".tech-camera-container");
const failureLink = document.querySelector(".tech-failure-link");

const videoInspector = new QRvideoInspector();

if (getParamValue("techCheck")) {
  videoContainerEl.style.display = "none";
  cameraFailureEl.style.display = "block";

  failureLink.href =
    localStorage.getItem("failureLink") || window.location.origin;
} else {
  videoContainerEl.style.display = "block";
}

failureLink.addEventListener("click", setContactToFailedList);

function getParamValue(param) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(param);
}

cameraCheckButtonEl.addEventListener("click", () => {
  videoInspector.inspect();
});

function getUserACId() {
  const segment = window.location.pathname.split("/").filter(Boolean)[0];
  if (!segment || segment === "admin") {
    return null;
  }
  return segment;
}

function techChecksPassed(user) {
  return user.camera && user.microphone && user.audio;
}

qrProcceedBtnEl.addEventListener("click", async () => {
  const userId = getUserACId();
  const response = await fetch(`/users/${userId}`);
  if (!response.ok) {
    qrFailureMessage.style.display = "block";
    return;
  }
  const user = await response.json();

  if (!techChecksPassed(user)) {
    qrFailureMessage.style.display = "block";
    return;
  }
  window.location.href = user.meetingLink;
});
