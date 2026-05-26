export class VideoInspector {
  #startVideoDetection() {
    navigator.mediaDevices
      .getUserMedia({ video: true })
      .then((stream) => {
        const video = document.querySelector("video");
        if ("srcObject" in video) {
          video.srcObject = stream;
        } else {
          video.src = window.URL.createObjectURL(stream);
        }
      })
      .catch((error) => {
        console.error("Error accessing camera:", error);
      });
  }

  handleCameraResult(result) {
    const videoContainerEl = document.querySelector(".tech-camera-container");
    const microContainerEl = document.querySelector(
      ".tech-microphone-container"
    );

    if (result) {
      videoContainerEl.style.display = "none";
      microContainerEl.style.display = "flex";
    }
  }

  getUserACId() {
    const currentUrl = window.location.href;
    const match = currentUrl.match(/\/(\w+)(?:\?.*)?$/);
    if (match) {
      return match[1];
    }
    return null;
  }

  inspect() {
    const cameraCheckBtn = document.querySelector("#camera-check-btn");
    cameraCheckBtn.disabled = true;
    const timeoutInSeconds = 25;

    Promise.all([
      faceapi.nets.tinyFaceDetector.loadFromUri("../../face-api-models/models"),
      faceapi.nets.faceLandmark68Net.loadFromUri(
        "../../face-api-models/models"
      ),
      faceapi.nets.faceRecognitionNet.loadFromUri(
        "../../face-api-models/models"
      ),
      faceapi.nets.faceExpressionNet.loadFromUri(
        "../../face-api-models/models"
      ),
    ]).then(this.#startVideoDetection);

    const intervalId = setInterval(async () => {
      const detections = await faceapi
        .detectAllFaces(video, new faceapi.TinyFaceDetectorOptions())
        .withFaceLandmarks()
        .withFaceExpressions();

      if (detections.length > 0) {
        clearInterval(intervalId);
        clearTimeout(timeoutId);

        this.handleCameraResult(true);
      }
    }, 2000);

    let reloadPage = false;

    const timeoutId = setTimeout(async () => {
      try {
        this.handleCameraResult(false);

        reloadPage = true;

        clearInterval(intervalId);
      } finally {
        if (reloadPage) {
          const currentURL = window.location.href;
          window.location.href = `${currentURL}&techCheck=failed`;
        }
      }
    }, timeoutInSeconds * 1000);
  }
}
