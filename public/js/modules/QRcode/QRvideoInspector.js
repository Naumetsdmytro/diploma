import { VideoInspector } from "../videoInspector.js";

export class QRvideoInspector extends VideoInspector {
  constructor() {
    super();
  }

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
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        checkResult,
      }),
    }).catch((error) => {
      console.error("Error sending camera check request:", error);
    });
  }

  getUserACId() {
    return super.getUserACId();
  }

  async updateUserResult() {
    const response = await fetch("/users");
    const users = await response.json();
    console.log(users);

    const userACId = this.getUserACId();
    const user = users.find((user) => user.id === userACId);

    await fetch(`users/${userACId}`, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        camera: true,
        microphone: false,
        audio: false,
        meetingLink: user.meetingLink,
        isPossibleToUsePhone: false,
      }),
    });
  }

  inspect() {
    super.inspect();
  }
}
