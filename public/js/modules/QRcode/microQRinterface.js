import { QRmicroInspector } from "./QRmicroInspector.js";
import { setContactToFailedList } from "../setContactToFailedList.js";

const microCheckButtonEl = document.getElementById("micro-check-btn");
const qrProcceedBtnEl = document.querySelector(".qr-button");
const qrFailureMessage = document.querySelector(".qr-failure-message");
const microBackdropEl = document.getElementById("microphone-check");
const microContainerEl = document.querySelector(".qr-microphone-container");
const microFailureEl = document.getElementById("microphone-failure-text");
const failureLink = document.querySelector(".tech-failure-link");

const microInspector = new QRmicroInspector();

if (getParamValue("techCheck")) {
  microContainerEl.style.display = "none";
  microFailureEl.style.display = "block";
  microBackdropEl.style.display = "none";

  failureLink.href =
    localStorage.getItem("failureLink") || window.location.origin;
} else {
  microContainerEl.style.display = "block";
}

failureLink.addEventListener("click", setContactToFailedList);

function getParamValue(param) {
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(param);
}

microCheckButtonEl.addEventListener("click", () => {
  microInspector.inspect();
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
