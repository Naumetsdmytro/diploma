const express = require("express");
const { Mutex } = require("async-mutex");
const { google } = require("googleapis");
const path = require("path");
const bodyParser = require("body-parser");
require("dotenv").config();
const http = require("http");
const socketIO = require("socket.io");
const users = require("./routes/users");
const { setDataSchema } = require("./schemas");
const { HttpError } = require("./helpers");
const passport = require("passport");
const { OAuth2Client } = require("google-auth-library");
const { Strategy: GoogleStrategy } = require("passport-google-oauth20");
const cors = require("cors");
const lock = new Mutex();
const app = express();
const server = http.createServer(app);
const io = socketIO(server);

// Set the static files directory
const publicDir = path.join(__dirname, "..", "public");

//  Middlewares --- USE
app.use(express.static(publicDir));
app.use(cors());
// COOP/COEP removed: they block cross-origin scripts (Socket.IO CDN, QRCode, annyang) with
// ERR_BLOCKED_BY_RESPONSE.NotSameOriginAfterDefaultedToSameOriginByCoep. Re-enable only if you
// need SharedArrayBuffer for face-api and serve all scripts same-origin.

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.json());
app.set("view engine", "ejs");
app.use("/users", users);

io.on("connection", (socket) => {
  socket.on("join", (userId) => {
    socket.join(userId);
  });
});

// .env Variables
const clientEmailE = process.env.CLIENT_EMAIL_E;
const privateKeyE = process.env.PRIVATE_KEY_E;
const clientEmailM = process.env.CLIENT_EMAIL_M;
const privateKeyM = process.env.PRIVATE_KEY_M;
const clientEmailTechCheck = process.env.CLIENT_EMAIL_TECH_CHECK;
const privateKeyTechCheck = process.env.PRIVATE_KEY_TECH_CHECK;
const spreadsheetId = process.env.SPREADSHEET_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const baselink = process.env.BASE_LINK;
const redirectUri = process.env.REDIRECT_URI;

const authE = new google.auth.JWT({
  email: clientEmailE,
  key: privateKeyE.replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const authM = new google.auth.JWT({
  email: clientEmailM,
  key: privateKeyM.replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const authTechCheck = new google.auth.JWT({
  email: clientEmailTechCheck,
  key: privateKeyTechCheck.replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// Cached Data "/getData"
let cachedData = null;
const userStates = {};

const dataArrayE = [];

const requestQueueE = [];
const requestQueueM = [];
const requestQueueTechCheck = [];

let isProcessingQueueE = false;
let isProcessingQueueM = false;
let isProcessingTechCheckQueue = false;

// Endpoints for Vlad to quickly restart server
app.get("/kapec", (req, res) => {
  cachedData = null;
  res.redirect(baselink);
});

app.get("/babita", (req, res) => {
  dataArrayE.length = 0;
  requestQueueE.length = 0;
  requestQueueM.length = 0;
  requestQueueTechCheck.length = 0;
  res.redirect(baselink);
});

// Get Date
app.get("/currentDateTime", (req, res) => {
  const currentDate = new Date();

  res.json({ currentDateTime: currentDate });
});

// Sign in with Google
const oauth2Client = new google.auth.OAuth2(
  clientId,
  clientSecret,
  redirectUri
);

app.get("/signin/google", (req, res) => {
  const { userId } = req.query;

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: "https://www.googleapis.com/auth/userinfo.profile",
    prompt: "select_account",
    redirect_uri: `${baselink}/oauth2callback`,
    state: userId,
  });
  res.redirect(authUrl);
});

app.get("/oauth2callback", async (req, res) => {
  const { code, state } = req.query;
  const userId = state;

  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const service = google.oauth2({ version: "v2", auth: oauth2Client });
    const response = await service.userinfo.get();
    const googleName = response.data.name;

    let redirectUrl = `${baselink}?signInSuccess=true&googleName=${googleName}`;
    if (userId) {
      redirectUrl = `${baselink}/${userId}?signInSuccess=true&googleName=${googleName}`;
    }

    res.redirect(redirectUrl);
  } catch (error) {
    console.error("Error:", error);
    res.status(500).send("An error occurred");
  }
});

// passport.use(
//   new GoogleStrategy(
//     {
//       clientID: clientId,
//       clientSecret,
//       redirectUri,
//     },
//     (token, tokenSecret, profile, done) => {
//       // Handle user authentication and create or retrieve user data
//       return done(null, profile);
//     }
//   )
// );

// passport.serializeUser((user, done) => {
//   done(null, user);
// });

// passport.deserializeUser((user, done) => {
//   done(null, user);
// });

// app.use(express.json());
// app.use(passport.initialize());

// app.get(
//   "/auth/google",
//   passport.authenticate("google", { scope: ["profile", "email"] })
// );

// app.get(
//   "/auth/google/callback",
//   passport.authenticate("google", { failureRedirect: "/" }),
//   (req, res) => {
//     // Successful authentication, redirect to the frontend or send a response
//     res.redirect("/");
//   }
// );

// app.post("/verifyGoogleToken", (req, res) => {
//   // Implement token verification logic here and send response
//   res.json({ success: true, userData: req.user });
// });

// Diploma / demo: allow joining the tech-check flow regardless of scheduled window
app.get("/isEduquestActive", (req, res) => {
  res.json({ data: [true, baselink] });
});

// Define a route to handle the /getDate request
app.get("/getData", (req, res) => {
  const currentDate = new Date();
  if (cachedData) {
    const startDate = new Date(cachedData[0][0]);
    if (new Date(startDate.getTime() + 40 * 60 * 1000) < currentDate) {
      cachedData = null;
      dataArrayE.length = 0;
      requestQueueE.length = 0;
      requestQueueM.length = 0;
      requestQueueTechCheck.length = 0;
    }
  }
  if (cachedData) {
    console.log("using cache");
    res.json({ data: cachedData });
    return;
  }
  console.log("not using cache");
  const sheets = google.sheets({ version: "v4", auth: authE });
  sheets.spreadsheets.values.get(
    {
      spreadsheetId: spreadsheetId,
      range: "Eduquests!A:G",
      majorDimension: "ROWS",
      valueRenderOption: "UNFORMATTED_VALUE",
    },
    async (err, response) => {
      if (err) {
        console.error("Error:", err);
        res.status(500).send("An error occurred");
        return;
      }

      const rows = response.data.values;
      if (rows && rows.length) {
        rows.shift();

        const closestDateArray = await getClosestDate(rows);
        const closestDate = new Date(closestDateArray[0]);
        let listOfInterns = [];
        if (
          new Date(closestDate.getTime() + 20 * 60 * 1000) >
          currentDate.getTime()
        ) {
          listOfInterns = await getInternsFromACbyListId(closestDateArray[5]);
        }
        cachedData = [[...closestDateArray, listOfInterns], rows];
        res.json({ data: cachedData });
      } else {
        res.json({ message: "No data found." });
      }
    }
  );
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

app.get("/", (req, res) => {
  res.render("mainPage");
});

async function getInternsFromACbyListId(listId) {
  let offset = 0;
  const contactsData = [];

  while (true) {
    const response = await fetch(
      `https://nobelcoaching22331.api-us1.com/api/3/contacts?listid=${listId}&limit=100&offset=${offset}&status=1`,
      {
        method: "GET",
        headers: {
          accept: "application/json",
          "Api-Token":
            "f787e136dd0dd5a2d1bc6e9e10b3e13e5e26cff03f7d3d31072e71a741df259e477736c6",
        },
      }
    );

    const result = await response.json();
    const contacts = await result.contacts;

    if (contacts.length > 0) {
      contactsData.push(...contacts);

      offset += 100;
    } else {
      break;
    }
  }

  return contactsData.map((data) => {
    return { id: data.id, email: data.email };
  });
}

async function getClosestDate(rows) {
  const currentDate = new Date();

  const closestDate = rows.reduce((closest, current) => {
    const currentRowDate = new Date(current[0]);
    const closestRowDate = new Date(closest[0]);

    const currentDiff = Math.abs(currentRowDate - currentDate);
    const closestDiff = Math.abs(closestRowDate - currentDate);

    return currentDiff < closestDiff &&
      currentRowDate.getTime() + 20 * 60 * 1000 >= currentDate
      ? current
      : closest;
  });
  return closestDate;
}

app.post("/cameraCheck/:userId", (req, res) => {
  const { userId } = req.params;
  const { checkResult } = req.body;

  checkResult
    ? io.to(userId).emit("cameraCheckPassed")
    : io.to(userId).emit("cameraCheckFailed");

  res.sendStatus(200);
});

app.post("/microphoneCheck/:userId", (req, res) => {
  const { userId } = req.params;
  const { checkResult } = req.body;

  checkResult
    ? io.to(userId).emit("microphoneCheckPassed")
    : io.to(userId).emit("microphoneCheckFailed");

  res.sendStatus(200);
});

app.post("/checkInternInFutureEQ", async (req, res) => {
  const internId = req.body.internId;
  const eduquestsList = cachedData[1];
  const closestEQ = cachedData[0];
  let isInternExist = false;

  const futureEQs = await getFutureEQs(eduquestsList, closestEQ);

  if (futureEQs.length > 0) {
    let isInternExist = false;

    for (const futureEQ of futureEQs) {
      const eqListId = futureEQ[5];
      const internData = await checkInternInEQById(internId, eqListId);

      if (internData.length > 0) {
        isInternExist = true;
        res.json({ futureEQ });
        return;
      }
    }

    if (!isInternExist) {
      res.json({ futureEQ: null });
    }
  } else {
    res.json({ futureEQ: null });
  }
});

async function checkInternInEQById(internId, listId) {
  const response = await fetch(
    `https://nobelcoaching22331.api-us1.com/api/3/contacts?ids=${internId}&listid=${listId}`,
    {
      method: "GET",
      headers: {
        accept: "application/json",
        "Api-Token":
          "f787e136dd0dd5a2d1bc6e9e10b3e13e5e26cff03f7d3d31072e71a741df259e477736c6",
      },
    }
  );

  const { contacts } = await response.json();
  return contacts;
}

async function getFutureEQs(eduquestsList, closestEQ) {
  return eduquestsList.filter((eduquest) => {
    const eduquestDateString = eduquest[0];
    const eduquestDate = new Date(eduquestDateString);
    const currentDate = new Date();
    if (
      eduquestDate.getTime() > currentDate.getTime() &&
      closestEQ[0] !== eduquestDateString &&
      !eduquest[4].includes("Backup")
    ) {
      return eduquest;
    }
  });
}

let roomCounter = 0;
function getNextRoomNumber(max) {
  roomCounter += 1;
  if (roomCounter > max) {
    roomCounter = 1;
  }
  return roomCounter;
}

// Example endpoint that uses the shared roomCounter.
app.post("/getNextRoomNumber", (req, res) => {
  const maxRooms = req.body.maxNumber;
  const roomNumber = getNextRoomNumber(maxRooms);
  res.json({ roomNumber });
});

app.post("/getEmailsFromEntered", (req, res) => {
  const email = req.body.data;
  let isEmailExist = false;
  let spreadsheetObj = dataArrayE[0];

  if (spreadsheetObj) {
    const enteredData = spreadsheetObj.data;
    isEmailExist = enteredData.find((data) => {
      if (data[0] === email) {
        return data;
      }
    });
  }

  if (isEmailExist) {
    res.json({ data: isEmailExist[3] });
  } else {
    res.json({ data: false });
  }
});

app.post("/setData", async (req, res) => {
  try {
    const data = req.body;

    const { error } = setDataSchema.validate(data);
    if (error) {
      throw HttpError(400, error.message);
    }

    // Acquire the lock to ensure exclusive access to the array
    await lock.acquire();

    if (data.sheetName === "Entered") {
      requestQueueE.push(data);
    } else if (data.sheetName === "Tech check") {
      requestQueueTechCheck.push(data);
    } else {
      requestQueueM.push(data);
    }
    // Release the lock
    lock.release();

    res.sendStatus(200);
    if (data.sheetName === "Entered") {
      if (!isProcessingQueueE) processQueueE();
    } else if (data.sheetName === "Tech check") {
      if (!isProcessingTechCheckQueue) processQueueTechCheck();
    } else {
      if (!isProcessingQueueM) processQueueM();
    }
  } catch (error) {
    const { status = 500, message = "Server error" } = error;
    res.status(status).json(message);
  }
});

async function processQueueE() {
  isProcessingQueueE = true;

  while (requestQueueE.length > 0) {
    const { spreadSheetId, data, sheetName } = requestQueueE[0]; // Retrieve the first request from the queue
    try {
      const sheets = google.sheets({ version: "v4", auth: authE });

      // Check if the spreadsheet object exists in the array
      let spreadsheetObj = dataArrayE.find(
        (obj) => obj.spreadSheetId === spreadSheetId
      );

      // If the spreadsheet object doesn't exist, create a new one with count = 2
      if (!spreadsheetObj) {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: spreadSheetId,
          range: `${sheetName}!A:B`,
          majorDimension: "ROWS",
        });
        const rows = response.data.values;
        const lastRow = rows.length + 1;
        spreadsheetObj = {
          spreadSheetId,
          sheets: [{ sheetName, count: lastRow }],
          data: [],
        };
        dataArrayE.push(spreadsheetObj);
      }

      let sheetObj = spreadsheetObj.sheets.find(
        (obj) => obj.sheetName === sheetName
      );

      spreadsheetObj.data.push(data);

      // If the sheet object doesn't exist, create a new one with count = 3
      if (!sheetObj) {
        spreadsheetObj.sheets.push({ sheetName, count: 3 });
        sheetObj = { sheetName, count: 2 };
      }

      const lastRow = sheetObj.count;

      // Set the data in the last row
      await sheets.spreadsheets.values.update({
        spreadsheetId: spreadSheetId,
        range: `${sheetName}!A${lastRow}:D${lastRow}`,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: [data],
        },
      });

      sheetObj.count++;
      requestQueueE.shift();
    } catch (error) {
      console.error(error);
      await new Promise((resolve) => setTimeout(resolve, 3000));
    }
  }
  isProcessingQueueE = false;
}

async function processQueueM() {
  isProcessingQueueM = true;

  while (requestQueueM.length > 0) {
    const { spreadSheetId, data, sheetName } = requestQueueM[0]; // Retrieve the first request from the queue
    try {
      const sheets = google.sheets({ version: "v4", auth: authM });

      // Get the last row in the sheet
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: spreadSheetId,
        range: `${sheetName}!A:C`,
        majorDimension: "ROWS",
      });
      const rows = response.data.values;
      const lastRow = rows.length + 1;

      // Set the data in the last row
      const range =
        data.length === 3
          ? `${sheetName}!A${lastRow}:C${lastRow}`
          : `${sheetName}!A${lastRow}:B${lastRow}`;

      await sheets.spreadsheets.values.update({
        spreadsheetId: spreadSheetId,
        range,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: [data],
        },
      });

      requestQueueM.shift();
    } catch (error) {
      console.error(error);
      await new Promise((resolve) => setTimeout(resolve, 3000));
    }
  }
  isProcessingQueueM = false;
}

async function processQueueTechCheck() {
  isProcessingTechCheckQueue = true;

  while (requestQueueTechCheck.length > 0) {
    const { spreadSheetId, data, sheetName } = requestQueueTechCheck[0]; // Retrieve the first request from the queue
    try {
      const sheets = google.sheets({ version: "v4", auth: authTechCheck });

      // Get the last row in the sheet
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: spreadSheetId,
        range: `${sheetName}!A:D`,
        majorDimension: "ROWS",
      });
      const rows = response.data.values;
      const lastRow = rows.length + 1;

      // Set the data in the last row
      await sheets.spreadsheets.values.update({
        spreadsheetId: spreadSheetId,
        range: `${sheetName}!A${lastRow}:D${lastRow}`,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: [data],
        },
      });

      requestQueueTechCheck.shift();
    } catch (error) {
      console.error(error);
      await new Promise((resolve) => setTimeout(resolve, 3000));
    }
  }
  isProcessingTechCheckQueue = false;
}

app.use((req, res) => {
  res.status(404).json({ message: "Not found" });
});

// Start the server
const port = process.env.PORT || 3000;
server.listen(port, () => {
  console.log(`Server is listening on port ${port}`);
});
