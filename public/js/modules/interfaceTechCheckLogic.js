import { MicroInspector } from "./microInspector.js";
import { VideoInspector } from "./videoInspector.js";
import { getParamValue } from "../main.js";

const microCheckButtonEl = document.getElementById("micro-check-btn");
const cameraCheckButtonEl = document.getElementById("camera-check-btn");
const cameraQREl = document.querySelector("#qr-camera");
const microphoneQREl = document.querySelector("#qr-microphone");
const cameraQRTextEl = document.querySelector("#qr-camera-text");
const microphoneQRTextEl = document.querySelector("#qr-microphone-text");
const deviceInfoText = document.querySelector(".device-info-text");
const contentContainer = document.querySelector(".content");
const signInButtonEl = document.querySelector(".signIn-google");

function startInspections() {
  console.log(
    "[interfaceTechCheckLogic] window load — wiring camera/mic (content stays visible)"
  );
  const isDesktop = performDeviceCheck();
  const cameraParam = getParamValue("camera");
  const microParam = getParamValue("microphone");
  const signInSuccess = getParamValue("signInSuccess");

  if (isDesktop && !cameraParam && !microParam) {
    if (!signInSuccess && signInButtonEl) {
      signInButtonEl.style.display = "inline-flex";
    }

    if (!cameraCheckButtonEl || !microCheckButtonEl) {
      console.error("Tech check buttons not found in the DOM");
      return;
    }

    const videoInspector = new VideoInspector();
    const microInspector = new MicroInspector();

    microCheckButtonEl.addEventListener("click", () => {
      microInspector.inspect();
    });

    cameraCheckButtonEl.addEventListener("click", () => {
      videoInspector.inspect();
    });

    cameraQREl.addEventListener("click", (evt) => {
      if (evt.target.id === "qr-camera") {
        cameraQREl.style.display = "none";
      }
    });

    microphoneQREl.addEventListener("click", (evt) => {
      if (evt.target.id === "qr-microphone") {
        microphoneQREl.style.display = "none";
      }
    });

    cameraQRTextEl.addEventListener("click", () => {
      cameraQREl.style.display = "flex";
    });

    microphoneQRTextEl.addEventListener("click", () => {
      microphoneQREl.style.display = "flex";
    });
  } else {
    contentContainer.style.display = "none";
    deviceInfoText.style.display = "block";
  }
}

function performDeviceCheck() {
  // Diploma / demo: always show .content (form + tech check). Old logic hid the whole UI on
  // touch devices and small screens after window load.
  return true;
}

window.addEventListener("load", startInspections);
