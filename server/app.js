const express = require("express");
const path = require("path");
const bodyParser = require("body-parser");
require("dotenv").config();
const http = require("http");
const socketIO = require("socket.io");
const cors = require("cors");
const { OAuth2Client } = require("google-auth-library");

const { connectDB } = require("./db/connect");
const AppSettings = require("./models/AppSettings");
const users = require("./routes/users");
const admin = require("./routes/admin");
const { HttpError } = require("./helpers");
const { recordTechCheck } = require("./services/techCheck");

const app = express();
const server = http.createServer(app);
const io = socketIO(server);

const publicDir = path.join(__dirname, "..", "public");
const baselink = process.env.BASE_LINK || "http://localhost:3000";
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const redirectUri = process.env.REDIRECT_URI;

app.use(express.static(publicDir));
app.use(cors());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.json());
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "..", "views"));

app.use("/users", users);
app.use("/admin", admin);

io.on("connection", (socket) => {
  socket.on("join", (userId) => {
    socket.join(userId);
  });
});

app.get("/api/config", async (req, res, next) => {
  try {
    const settings = await AppSettings.getSettings();
    res.json({
      meetingLink: settings.meetingLink,
      failureLink: settings.failureLink,
    });
  } catch (error) {
    next(error);
  }
});

app.get("/admin", async (req, res, next) => {
  try {
    const settings = await AppSettings.getSettings();
    res.render("adminPage", {
      meetingLink: settings.meetingLink,
      failureLink: settings.failureLink,
    });
  } catch (error) {
    next(error);
  }
});

const oauth2Client = new OAuth2Client(clientId, clientSecret, redirectUri);

app.get("/signin/google", (req, res) => {
  const { userId } = req.query;

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: "https://www.googleapis.com/auth/userinfo.profile",
    prompt: "select_account",
    redirect_uri: `${baselink}/oauth2callback`,
    state: userId || "",
  });
  res.redirect(authUrl);
});

app.get("/oauth2callback", async (req, res) => {
  const { code, state } = req.query;
  const userId = state;

  try {
    const { tokens } = await oauth2Client.getToken(code);
    const accessToken = tokens.access_token;

    const profileResponse = await fetch(
      "https://www.googleapis.com/oauth2/v2/userinfo",
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );
    const profile = await profileResponse.json();
    const googleName = profile.name || "";

    let redirectUrl = `${baselink}?signInSuccess=true&googleName=${encodeURIComponent(googleName)}`;
    if (userId) {
      redirectUrl = `${baselink}/${userId}?signInSuccess=true&googleName=${encodeURIComponent(googleName)}`;
    }

    res.redirect(redirectUrl);
  } catch (error) {
    console.error("OAuth error:", error);
    res.status(500).send("An error occurred during Google sign-in");
  }
});

app.post("/cameraCheck/:userId", async (req, res, next) => {
  try {
    const { userId } = req.params;
    const { checkResult } = req.body;

    if (checkResult) {
      await recordTechCheck(userId, "camera", {
        passed: true,
        device: "mobile",
      });
      io.to(userId).emit("cameraCheckPassed");
    } else {
      await recordTechCheck(userId, "camera", {
        passed: false,
        device: "mobile",
      });
      io.to(userId).emit("cameraCheckFailed");
    }

    res.sendStatus(200);
  } catch (error) {
    next(error);
  }
});

app.post("/microphoneCheck/:userId", async (req, res, next) => {
  try {
    const { userId } = req.params;
    const { checkResult } = req.body;

    if (checkResult) {
      await recordTechCheck(userId, "microphone", {
        passed: true,
        device: "mobile",
      });
      io.to(userId).emit("microphoneCheckPassed");
    } else {
      await recordTechCheck(userId, "microphone", {
        passed: false,
        device: "mobile",
      });
      io.to(userId).emit("microphoneCheckFailed");
    }

    res.sendStatus(200);
  } catch (error) {
    next(error);
  }
});

app.get("/", (req, res) => {
  res.render("mainPage");
});

app.get("/:id", (req, res) => {
  const cameraParameterValue = req.query.camera;
  const microParameterValue = req.query.microphone;

  if (cameraParameterValue) {
    res.render("cameraQrPage");
  } else if (microParameterValue) {
    res.render("microphoneQrPage");
  } else {
    res.render("mainPage");
  }
});

app.use((err, req, res, next) => {
  const status = err.status || 500;
  res.status(status).json({ message: err.message || "Server error" });
});

app.use((req, res) => {
  res.status(404).json({ message: "Not found" });
});

const port = process.env.PORT || 3000;

connectDB()
  .then(() => AppSettings.getSettings())
  .then(() => {
    server.listen(port, () => {
      console.log(`Server is listening on port ${port}`);
    });
  })
  .catch((error) => {
    console.error("Failed to start server:", error);
    process.exit(1);
  });
