/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */

import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

// MSAL Configuration
const msalConfig = {
  auth: {
    clientId: "73f655a9-0321-4ca5-b13a-bcb7759d4679",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://dublicate-outlook-addin.vercel.app/commands.html",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false, // Recommended for Office Add-ins
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

// Login request configuration
const loginRequest = {
  scopes: ["User.Read", "Calendars.ReadWrite"],
};

// Track initialization promise
let msalInitialized = false;

// Simplified initialization flow
async function initializeMSAL() {
  if (!msalInitialized) {
    try {
      await msalInstance.initialize();
      console.log("MSAL initialized successfully");
      msalInitialized = true;
    } catch (error) {
      console.error("MSAL initialization failed:", error);
      throw error;
    }
  }
  return msalInstance;
}

// Track authentication state
let authState = {
  inProgress: false,
  retryCount: 0,
  MAX_RETRIES: 2,
};

// Reset authentication state
function resetAuthState() {
  authState = {
    inProgress: false,
    retryCount: 0,
    MAX_RETRIES: 2,
  };
}

async function getAccessToken() {
  await initializeMSAL();

  if (authState.retryCount >= authState.MAX_RETRIES) {
    resetAuthState();
    throw new Error("Maximum authentication attempts reached");
  }

  try {
    if (authState.inProgress) {
      throw new Error("Authentication already in progress");
    }

    authState.inProgress = true;
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
      console.log("Attempting login popup");
      const response = await msalInstance.loginPopup(loginRequest);
      console.log("Login response:", response); // Log full response
      if (!response.accessToken) {
        throw new Error("No access token in login response");
      }
      return response.accessToken;
    }

    try {
      console.log("Attempting silent token acquisition");
      const result = await msalInstance.acquireTokenSilent({
        // Fixed typo here
        ...loginRequest,
        account: accounts[0],
      });
      console.log("Silent acquisition result:", result);
      return result.accessToken;
    } catch (silentError) {
      if (silentError instanceof InteractionRequiredAuthError) {
        console.log("Falling back to interactive acquisition");
        const result = await msalInstance.acquireTokenPopup({
          ...loginRequest,
          account: accounts[0],
        });
        return result.accessToken;
      }
      throw silentError;
    }
  } catch (error) {
    console.error("Token acquisition error:", error);
    if (error.errorCode === "interaction_in_progress") {
      authState.retryCount++;
      await new Promise((resolve) => setTimeout(resolve, 1000 * authState.retryCount));
      return getAccessToken();
    }
    throw error;
  } finally {
    authState.inProgress = false;
  }
}

const getAsyncValue = (property) => {
  let item = Office.context.mailbox.item;
  return new Promise((resolve, reject) => {
    // Check if the platform is desktop; on web, these properties are available directly
    if (Office.context.platform === "OfficeOnline") {
      resolve(item[property]); // Direct access for web version
    } else {
      item[property].getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(`Error retrieving ${property}: ${result.error}`);
        }
      });
    }
  });
};

async function action(event) {
  try {
    console.log("Action called");
    const token = await getAccessToken();
    console.log("Access Token:", token);

    ///////////////////////////////////////////////////
    let item = Office.context.mailbox.item;
    // Fetch properties with different methods depending on the platform
    const subject = await getAsyncValue("subject");
    const start = await getAsyncValue("start");
    const end = await getAsyncValue("end");
    const location = await getAsyncValue("location");

    // Fetch the body (both web and desktop use getAsync for the body)
    const body = await new Promise((resolve, reject) => {
      item.body.getAsync("html", (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(`Error retrieving event description: ${result.error}`);
        }
      });
    });

    // Fetch attendees (both required and optional)
    const getAttendeesAsync = (attendeeType) => {
      return new Promise((resolve, reject) => {
        // Check if the platform is desktop; on web, these properties are available directly
        if (Office.context.platform === "OfficeOnline") {
          resolve(item[attendeeType]); // Direct access for web version
          return;
        } else {
          item[attendeeType].getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value);
            } else {
              reject(`Error retrieving ${attendeeType}: ${result.error}`);
            }
          });
        }
      });
    };

    const requiredAttendees = await getAttendeesAsync("requiredAttendees");
    const optionalAttendees = await getAttendeesAsync("optionalAttendees");

    const newEvent = {};

    // Only include subject if it exists
    if (subject) newEvent.subject = subject;

    // Only include start if it exists
    if (start) {
      newEvent.start = {
        dateTime: start.toISOString(),
        timeZone: "UTC",
      };
    }

    // Only include end if it exists
    if (end) {
      newEvent.end = {
        dateTime: end.toISOString(),
        timeZone: "UTC",
      };
    }

    // Only include location if it exists
    if (location) {
      newEvent.location = { displayName: location };
    }

    // Only include body if it exists
    if (body) {
      newEvent.body = { contentType: "HTML", content: body };
    }

    newEvent.attendees = [];

    // Map required attendees (if they exist)
    if (requiredAttendees && requiredAttendees.length > 0) {
      const required = requiredAttendees.map((attendee) => ({
        emailAddress: {
          address: attendee.emailAddress, // Use "address" instead of "emailAddress"
          name: attendee.displayName,
        },
        type: "required", // Explicitly set the type
      }));
      newEvent.attendees.push(...required);
    }

    // Map optional attendees (if they exist)
    if (optionalAttendees && optionalAttendees.length > 0) {
      const optional = optionalAttendees.map((attendee) => ({
        emailAddress: {
          address: attendee.emailAddress,
          name: attendee.displayName,
        },
        type: "optional", // Explicitly set the type
      }));
      newEvent.attendees.push(...optional);
    }

    console.log(newEvent);
    ///////////////////////////////////////////////////

    await createCalendarEvent(newEvent, event, token);
    // await testMeEndpoint(token);
    event.completed();
  } catch (error) {
    console.error("Action failed:", error);

    resetAuthState();
    event.completed();
  }
}

async function testMeEndpoint(token) {
  const graphEndpoint = "https://graph.microsoft.com/v1.0/me";
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  try {
    const response = await fetch(graphEndpoint, {
      method: "GET",
      headers: headers,
    });

    const data = await response.json();

    if (!response.ok) {
      console.error("❌ Error fetching /me endpoint:", data);
      return null;
    }

    console.log("✅ /me response:", data);
    return data; // Returns user profile data
  } catch (error) {
    console.error("🚨 Network/request error:", error);
    return null;
  }
}

async function createCalendarEvent(eventData, event, token) {
  const graphEndpoint = "https://graph.microsoft.com/v1.0/me/events";
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  try {
    const response = await fetch(graphEndpoint, {
      method: "POST",
      headers: headers,
      body: JSON.stringify(eventData),
    });

    // Log HTTP status and response for debugging
    console.log("HTTP Status:", response.status);
    const data = await response.json();

    if (!response.ok) {
      // Handle Microsoft Graph API errors (e.g., 4xx/5xx responses)
      throw new Error(`API Error: ${JSON.stringify(data)}`);
    }

    console.log("Event created successfully:", data);
    event.completed();
  } catch (error) {
    console.error("Error creating event:", error);
    event.completed();
  }
}

// Register the function with Office.
Office.actions.associate("action", action);
