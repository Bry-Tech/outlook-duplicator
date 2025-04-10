import React from "react";

export default function App() {
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

  const getEventData = async () => {
    try {
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

      // Only include attendees if they exist
      if (requiredAttendees && requiredAttendees.length > 0) {
        newEvent.requiredAttendees = requiredAttendees.map((attendee) => ({
          emailAddress: attendee.emailAddress,
          name: attendee.displayName,
        }));
      }

      if (optionalAttendees && optionalAttendees.length > 0) {
        newEvent.optionalAttendees = optionalAttendees.map((attendee) => ({
          emailAddress: attendee.emailAddress,
          name: attendee.displayName,
        }));
      }

      console.log(newEvent);
    } catch (error) {
      console.error(error);
    }
  };

  return (
    <div>
      <button onClick={getEventData}>getEventData</button>
    </div>
  );
}
