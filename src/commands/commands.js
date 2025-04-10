/* global Office */

Office.onReady(() => {
  console.log("Office.js is ready");
  Office.actions.associate("action", handleAction);
});

let loginDialog;

// 🔄 Main action
async function handleAction(event) {
  try {
    console.log("Action triggered");

    const token = await getValidToken();
    // showNotification(event, "Working on it...");
    if (!token) {
      console.error("No valid token found.");
      // showNotification(event, "No valid token found.");
      event.completed();
      return;
    }

    const item = Office.context.mailbox.item;

    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      const subject = item.subject;
      const start = item.start;
      const end = item.end;
      const location = item.location;

      item.body.getAsync("html", async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const body = result.value;

          // showNotification(event, "start");
          const newEvent = {
            subject,
            start: { dateTime: start.toISOString(), timeZone: "UTC" },
            end: { dateTime: end.toISOString(), timeZone: "UTC" },
            location: { displayName: location },
            body: { contentType: "HTML", content: body },
          };
          // showNotification(event, "end");

          // showNotification(event, "Inside...");
          await createCalendarEvent(newEvent, token);
          console.log("Event created successfully!");
          // showNotification(event, "Event created successfully!");
        } else {
          console.error("Failed to get body content:", result.error.message);
        }

        event.completed();
      });
    } else {
      console.log("This is not a calendar item.");
      event.completed();
    }
  } catch (error) {
    console.error("Error in action:", error);
    event.completed();
  }
}

// ✅ Get valid token or login again
async function getValidToken() {
  const stored = localStorage.getItem("tokenData");
  if (stored) {
    const { token, expiresAt } = JSON.parse(stored);
    if (Date.now() < expiresAt) {
      return token;
    }
  }

  return await promptLoginDialog();
}

// 🔐 Open auth dialog and store token
function promptLoginDialog() {
  return new Promise((resolve, reject) => {
    const clientId = "c87c26dc-39f9-48c4-bfa0-e638588abb5f";
    const redirect = "https://sebastian-outlook-addin.vercel.app/assets/login.html";
    const scopes = "Calendars.ReadWrite";

    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${encodeURIComponent(redirect)}&scope=${encodeURIComponent(scopes)}&response_mode=fragment&state=12345&nonce=678910`;

    Office.context.ui.displayDialogAsync(authUrl, { height: 50, width: 30 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to open login dialog");
        reject("Dialog failed");
        return;
      }

      loginDialog = result.value;
      loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const token = arg.message;
        if (token && typeof token === "string" && token.length > 0) {
          const expiresAt = Date.now() + 3600 * 1000; // Assuming 1 hour token life
          localStorage.setItem("tokenData", JSON.stringify({ token, expiresAt }));
          loginDialog.close();
          resolve(token);
        } else {
          loginDialog.close();
          reject("Invalid token received");
        }
      });
    });
  });
}

// 📅 Send request to Microsoft Graph
async function createCalendarEvent(eventData, token) {
  const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(eventData),
  });

  if (!response.ok) {
    const errorDetails = await response.json();
    console.error("Graph API error:", errorDetails);
    throw new Error("Failed to create event");
  }

  return await response.json();
}

function showNotification(event, msg) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

// /*
//  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
//  * See LICENSE in the project root for license information.
//  */

// /* global Office */

// /**
//  * Shows a notification when the add-in command is executed.
//  * @param event {Office.AddinCommands.Event}
//  */

// import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

// // MSAL Configuration
// const msalConfig = {
//   auth: {
//     clientId: "c87c26dc-39f9-48c4-bfa0-e638588abb5f",
//     authority: "https://login.microsoftonline.com/common",
//     redirectUri: "https://localhost:3000/commands.html",
//   },
//   cache: {
//     cacheLocation: "localStorage",
//     storeAuthStateInCookie: false, // Recommended for Office Add-ins
//   },
// };

// const msalInstance = new PublicClientApplication(msalConfig);

// // Login request configuration
// const loginRequest = {
//   scopes: ["User.Read", "Calendars.ReadWrite"],
// };

// // Track initialization promise
// let msalInitialized = false;

// // Simplified initialization flow
// async function initializeMSAL() {
//   if (!msalInitialized) {
//     try {
//       await msalInstance.initialize();
//       console.log("MSAL initialized successfully");
//       msalInitialized = true;
//     } catch (error) {
//       console.error("MSAL initialization failed:", error);
//       throw error;
//     }
//   }
//   return msalInstance;
// }

// // Track authentication state
// let authState = {
//   inProgress: false,
//   retryCount: 0,
//   MAX_RETRIES: 2,
// };

// // Reset authentication state
// function resetAuthState() {
//   authState = {
//     inProgress: false,
//     retryCount: 0,
//     MAX_RETRIES: 2,
//   };
// }

// async function getAccessToken() {
//   await initializeMSAL();

//   if (authState.retryCount >= authState.MAX_RETRIES) {
//     resetAuthState();
//     throw new Error("Maximum authentication attempts reached");
//   }

//   try {
//     if (authState.inProgress) {
//       throw new Error("Authentication already in progress");
//     }

//     authState.inProgress = true;
//     const accounts = msalInstance.getAllAccounts();

//     if (accounts.length === 0) {
//       console.log("Attempting login popup");
//       const response = await msalInstance.loginPopup(loginRequest);
//       console.log("Login response:", response); // Log full response
//       if (!response.accessToken) {
//         throw new Error("No access token in login response");
//       }
//       return response.accessToken;
//     }

//     try {
//       console.log("Attempting silent token acquisition");
//       const result = await msalInstance.acquireTokenSilent({
//         // Fixed typo here
//         ...loginRequest,
//         account: accounts[0],
//       });
//       console.log("Silent acquisition result:", result);
//       return result.accessToken;
//     } catch (silentError) {
//       if (silentError instanceof InteractionRequiredAuthError) {
//         console.log("Falling back to interactive acquisition");
//         const result = await msalInstance.acquireTokenPopup({
//           ...loginRequest,
//           account: accounts[0],
//         });
//         return result.accessToken;
//       }
//       throw silentError;
//     }
//   } catch (error) {
//     console.error("Token acquisition error:", error);
//     if (error.errorCode === "interaction_in_progress") {
//       authState.retryCount++;
//       await new Promise((resolve) => setTimeout(resolve, 1000 * authState.retryCount));
//       return getAccessToken();
//     }
//     throw error;
//   } finally {
//     authState.inProgress = false;
//   }
// }

// // Updated action function
// async function action(event) {
//   try {
//     console.log("Action called");
//     const token = await getAccessToken();
//     console.log("Access Token:", token);

//     var item = Office.context.mailbox.item;

//     if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
//       // Get relevant details from the active calendar event
//       var subject = item.subject;
//       var start = item.start;
//       var end = item.end;
//       var location = item.location;

//       // Use getAsync() to fetch the body content as it is an ItemBody object
//       item.body.getAsync("html", function (result) {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           var body = result.value;
//           console.log("Body: " + body);

//           // Now, create a new calendar event using the details from the current event
//           var newEvent = {
//             subject: subject,
//             start: {
//               dateTime: start.toISOString(),
//               timeZone: "UTC",
//             },
//             end: {
//               dateTime: end.toISOString(),
//               timeZone: "UTC",
//             },
//             location: {
//               displayName: location,
//             },
//             body: {
//               contentType: "HTML",
//               content: body,
//             },
//           };

//           // Make a POST request to Microsoft Graph to create the new event
//           createCalendarEvent(newEvent, event, token);
//         } else {
//           console.error("Failed to get body content: " + result.error.message);
//           event.completed();
//         }
//       });
//     } else {
//       console.log("This is not a calendar event.");
//       event.completed();
//     }
//   } catch (error) {
//     console.error("Action failed:", error);

//     resetAuthState();
//     event.completed();
//   }
// }

// function createCalendarEvent(eventData, event, token) {
//   // Microsoft Graph API endpoint to create an event
//   var graphEndpoint = "https://graph.microsoft.com/v1.0/me/events";

//   // Set up the request headers with the token
//   var headers = new Headers({
//     Authorization: "Bearer " + token,
//     "Content-Type": "application/json",
//   });

//   // Make the POST request to create the event
//   fetch(graphEndpoint, {
//     method: "POST",
//     headers: headers,
//     body: JSON.stringify(eventData),
//   })
//     .then((response) => response.json())
//     .then((data) => {
//       console.log("New event created successfully!");
//       event.completed();
//     })
//     .catch((error) => {
//       console.error("Error creating event:", error);
//       event.completed();
//     });
// }

// Office.onReady(() => {
//   console.log("Office.js is ready");
//   Office.actions.associate("action", action);
// });
