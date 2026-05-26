export class MicroInspector {
  #matchingWord;

  constructor(matchingWord = "") {
    this.#matchingWord = matchingWord;
    this.messageDisplayed = false;
  }

  #findMatchesInSpeechResults(speechResults) {
    for (const result of Array.from(speechResults)) {
      const transcript = result[0].transcript.trim().toLowerCase();

      if (transcript.trim() !== "") return true;
    }
    return false;
  }

  handleMicroResult(result) {
    const microContainerEl = document.querySelector(
      ".tech-microphone-container"
    );
    const audioContainerEl = document.querySelector(".tech-audio-container");
    const issuesAudioLink = document.querySelector(".issues-audio-link");

    if (result) {
      audioContainerEl.style.display = "flex";
      microContainerEl.style.display = "none";
      this.messageDisplayed = true;
      issuesAudioLink.href = localStorage.getItem("failureLink");
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
    const timeoutInSeconds = 25;
    let recognition;

    if ("SpeechRecognition" in window) {
      recognition = new window.SpeechRecognition();
    } else if ("webkitSpeechRecognition" in window) {
      recognition = new window.webkitSpeechRecognition();
    } else {
      console.log("Speech recognition not supported in this browser");
    }

    const microphoneBackdropEl = document.getElementById("microphone-check");
    microphoneBackdropEl.style.display = "flex";

    recognition.lang = "en-US";
    recognition.interimResults = true;
    let recognitionTimeout;

    recognition.addEventListener("result", (e) => {
      if (!this.messageDisplayed) {
        clearTimeout(recognitionTimeout);
        const speechRecognitionResults = e.results;
        const findMatchesResult = this.#findMatchesInSpeechResults(
          speechRecognitionResults
        );

        recognitionTimeout = setTimeout(() => {
          if (findMatchesResult) {
            clearTimeout(timeoutId);
            this.handleMicroResult(true);
          }
        }, 9000);
      }
    });

    recognition.start();

    let reloadPage = false;

    const timeoutId = setTimeout(async () => {
      try {
        this.handleMicroResult(false);

        reloadPage = true;
      } catch (error) {
        console.error("Error in setTimeout callback:", error);
      } finally {
        if (reloadPage) {
          const currentURL = window.location.href;
          window.location.href = `${currentURL}&techCheck=failed`;
        }
      }
    }, timeoutInSeconds * 1000);
  }
}
