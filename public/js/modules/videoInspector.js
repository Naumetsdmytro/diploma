const MODEL_URL = "/face-api-models/models";

export class VideoInspector {
  getUserACId() {
    const segment = window.location.pathname.split("/").filter(Boolean)[0];
    if (!segment || segment === "admin") {
      return null;
    }
    return segment;
  }

  handleCameraResult(result) {
    const videoContainerEl = document.querySelector(".tech-camera-container");
    const microContainerEl = document.querySelector(
      ".tech-microphone-container"
    );

    if (result) {
      this.#persistCheck("camera", true);
      videoContainerEl.style.display = "none";
      microContainerEl.style.display = "flex";
    } else {
      this.#persistCheck("camera", false);
    }
  }

  #persistCheck(step, passed) {
    const userId = this.getUserACId();
    if (!userId) return;

    fetch(`/users/${userId}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ [step]: passed }),
    }).catch((error) => console.error(`Failed to save ${step} result:`, error));
  }

  async inspect() {
    const cameraCheckBtn = document.querySelector("#camera-check-btn");
    const videoEl = document.getElementById("video");
    const faceapi = window.faceapi;
    const originalLabel = cameraCheckBtn?.textContent || "Check my camera";

    if (!cameraCheckBtn) {
      console.error("Camera check button not found");
      return;
    }
    if (!videoEl) {
      console.error("Video element not found");
      return;
    }
    if (!faceapi) {
      console.error("face-api.js is not loaded — refresh the page");
      return;
    }
    if (!navigator.mediaDevices?.getUserMedia) {
      console.error("Camera is not supported in this browser");
      return;
    }

    cameraCheckBtn.disabled = true;

    const timeoutInSeconds = 25;
    let intervalId;
    let timeoutId;
    let stream;

    const cleanup = () => {
      clearInterval(intervalId);
      clearTimeout(timeoutId);
    };

    const resetUi = () => {
      cleanup();
      cameraCheckBtn.disabled = false;
      cameraCheckBtn.textContent = originalLabel;
    };

    try {
      cameraCheckBtn.textContent = "Allow camera access...";
      stream = await navigator.mediaDevices.getUserMedia({ video: true });
      videoEl.srcObject = stream;
      await videoEl.play();

      cameraCheckBtn.textContent = "Loading face detection...";
      await faceapi.nets.tinyFaceDetector.loadFromUri(MODEL_URL);

      cameraCheckBtn.textContent = "Look at the camera...";

      intervalId = setInterval(async () => {
        try {
          const detections = await faceapi.detectAllFaces(
            videoEl,
            new faceapi.TinyFaceDetectorOptions()
          );

          if (detections.length > 0) {
            cleanup();
            if (stream) {
              stream.getTracks().forEach((track) => track.stop());
            }
            this.handleCameraResult(true);
          }
        } catch (error) {
          console.error("Face detection error:", error);
        }
      }, 2000);

      timeoutId = setTimeout(() => {
        cleanup();
        if (stream) {
          stream.getTracks().forEach((track) => track.stop());
        }
        this.handleCameraResult(false);
        window.location.href = `${window.location.href}&techCheck=failed`;
      }, timeoutInSeconds * 1000);
    } catch (error) {
      resetUi();
      if (stream) {
        stream.getTracks().forEach((track) => track.stop());
      }
      console.error("Camera check failed:", error);
      cameraCheckBtn.textContent = "Camera blocked — try again";
      setTimeout(() => {
        cameraCheckBtn.textContent = originalLabel;
      }, 3000);
    }
  }
}
