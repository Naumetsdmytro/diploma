import { VideoInspector } from "../videoInspector.js";

export class QRvideoInspector extends VideoInspector {
  handleCameraResult(result) {
    const videoContainerEl = document.querySelector(".tech-camera-container");
    const resultContainerEl = document.querySelector(".qr-result");

    if (result) {
      videoContainerEl.style.display = "none";
      resultContainerEl.style.display = "block";
      this.sendCameraCheckSignal(this.getUserACId(), true);
    } else {
      this.sendCameraCheckSignal(this.getUserACId(), false);
    }
  }

  async sendCameraCheckSignal(userId, checkResult) {
    fetch(`/cameraCheck/${userId}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ checkResult }),
    }).catch((error) => {
      console.error("Error sending camera check request:", error);
    });
  }

  inspect() {
    super.inspect();
  }
}
