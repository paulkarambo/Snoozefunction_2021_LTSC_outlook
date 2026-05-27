// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.

/* global document, Office, OfficeExtension, OfficeRuntime, flatpickr */

// Initialize the add-in
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
    
    // Initialize Flatpickr for custom date/time selection
    const customTimeInput = document.getElementById("custom-snooze-time") as HTMLInputElement;
    flatpickr(customTimeInput, {
      enableTime: true,
      dateFormat: "Y-m-d H:i",
      minDate: "today",
      time_24hr: true
    });
    
    // Show/hide custom time input based on radio selection
    const customRadio = document.getElementById("custom") as HTMLInputElement;
    customRadio.addEventListener("change", () => {
      const customTimeDiv = document.getElementById("custom-date-time");
      if (customRadio.checked && customTimeDiv) {
        customTimeDiv.style.display = "block";
      } else if (customTimeDiv) {
        customTimeDiv.style.display = "none";
      }
    });
    
    // Handle snooze button click
    document.getElementById("snooze-button")?.addEventListener("click", snoozeEmail);
    document.getElementById("cancel-button")?.addEventListener("click", closeDialog);
  }
});

// Close the dialog
function closeDialog() {
  Office.context.ui.closeContainer();
}

// Snooze the selected email
async function snoozeEmail() {
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      showNotification("Error", "No email selected");
      return;
    }
    
    // Get selected snooze duration
    const selectedDuration = document.querySelector("input[name='snooze-duration']:checked") as HTMLInputElement;
    if (!selectedDuration) {
      showNotification("Error", "Please select a snooze duration");
      return;
    }
    
    // Calculate return time
    let returnTime: Date;
    switch (selectedDuration.value) {
      case "later-today":
        returnTime = new Date();
        returnTime.setHours(16, 0, 0, 0);
        if (returnTime <= new Date()) {
          returnTime.setDate(returnTime.getDate() + 1);
        }
        break;
      case "tomorrow":
        returnTime = new Date();
        returnTime.setDate(returnTime.getDate() + 1);
        returnTime.setHours(8, 0, 0, 0);
        break;
      case "this-weekend":
        returnTime = new Date();
        const dayOfWeek = returnTime.getDay();
        const daysToSaturday = dayOfWeek <= 6 ? 6 - dayOfWeek : 0;
        returnTime.setDate(returnTime.getDate() + daysToSaturday);
        returnTime.setHours(9, 0, 0, 0);
        break;
      case "next-week":
        returnTime = new Date();
        returnTime.setDate(returnTime.getDate() + 7);
        returnTime.setHours(8, 0, 0, 0);
        break;
      case "custom":
        const customTimeInput = document.getElementById("custom-snooze-time") as HTMLInputElement;
        const customTime = customTimeInput.value;
        if (!customTime) {
          showNotification("Error", "Please select a custom date/time");
          return;
        }
        returnTime = new Date(customTime);
        break;
      default:
        showNotification("Error", "Invalid snooze duration");
        return;
    }
    
    // Get snooze action
    const moveBack = (document.getElementById("move-back") as HTMLInputElement).checked;
    const flagEmail = (document.getElementById("flag-email") as HTMLInputElement).checked;
    
    // Store snooze metadata
    const snoozeData = {
      returnTime: returnTime.getTime(),
      moveBack,
      flagEmail,
      originalItemId: item.itemId
    };
    
    // Store in localStorage
    const snoozedItems = JSON.parse(localStorage.getItem("snoozedItems") || "[]");
    snoozedItems.push(snoozeData);
    localStorage.setItem("snoozedItems", JSON.stringify(snoozedItems));
    
    // Move email to the "Snoozed" folder
    const folderId = await getSnoozeFolderId();
    if (!folderId) {
      showNotification("Error", "Could not find or create 'Snoozed' folder");
      return;
    }
    
    // Move the item
    await item.move(folderId);
    
    // Close the dialog
    Office.context.ui.closeContainer();
    
    // Show confirmation
    showNotification("Success", `Email snoozed until ${returnTime.toLocaleString()}`);
  } catch (error) {
    showNotification("Error", error instanceof Error ? error.message : "Unknown error");
  }
}

// Get or create the "Snoozed" folder
async function getSnoozeFolderId(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const ewsId = Office.context.mailbox.convertToEwsId(
          Office.context.mailbox.userProfile.emailAddress,
          Office.MailboxEnums.RestVersion.v2_0
        );
        
        // Search for existing "Snoozed" folder
        Office.context.mailbox.makeEwsRequestAsync(
          `<FindFolder Traversal="Shallow">
            <FolderShape>
              <t:BaseShape>Default</t:BaseShape>
            </FolderShape>
            <ParentFolderIds>
              <DistinguishedFolderId Id="msgfolderroot"/>
            </ParentFolderIds>
            <Restriction>
              <IsEqualTo>
                <FieldURI FieldURI="folder:DisplayName"/>
                <FieldURIOrConstant>
                  <Constant Value="Snoozed"/>
                </FieldURIOrConstant>
              </IsEqualTo>
            </Restriction>
          </FindFolder>`,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              const response = asyncResult.value;
              if (response.includes('<t:FolderId Id=')) {
                const idMatch = response.match(/<t:FolderId Id="([^"]+)"/);
                if (idMatch && idMatch[1]) {
                  resolve(idMatch[1]);
                  return;
                }
              }
              // Create folder if it doesn't exist
              Office.context.mailbox.makeEwsRequestAsync(
                `<CreateFolder>
                  <ParentFolderId>
                    <DistinguishedFolderId Id="msgfolderroot"/>
                  </ParentFolderId>
                  <Folders>
                    <Folder>
                      <DisplayName>Snoozed</DisplayName>
                    </Folder>
                  </Folders>
                </CreateFolder>`,
                (createResult) => {
                  if (createResult.status === Office.AsyncResultStatus.Succeeded) {
                    const createResponse = createResult.value;
                    const idMatch = createResponse.match(/<t:FolderId Id="([^"]+)"/);
                    if (idMatch && idMatch[1]) {
                      resolve(idMatch[1]);
                    } else {
                      reject(new Error('Failed to get folder ID after creation'));
                    }
                  } else {
                    reject(new Error('Failed to create snooze folder'));
                  }
                }
              );
            } else {
              reject(new Error('Failed to search for snooze folder'));
            }
          }
        );
      } else {
        reject(new Error('Failed to get callback token'));
      }
    });
  });
}

// Show a notification
function showNotification(title: string, message: string): void {
  const details: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: `${title}: ${message}`,
    icon: "Icon.80x80",
    persistent: true,
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync(`snoozeNotification-${title}`, details);
}