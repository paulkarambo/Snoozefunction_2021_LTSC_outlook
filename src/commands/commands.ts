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
 * @param event
 */
let snoozeFolderId: string | null = null;

// Initialize the snooze folder
async function initializeSnoozeFolder(): Promise<void> {
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
                // Extract folder ID
                const idMatch = response.match(/<t:FolderId Id="([^"]+)"/);
                if (idMatch && idMatch[1]) {
                  snoozeFolderId = idMatch[1];
                  resolve();
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
                      snoozeFolderId = idMatch[1];
                    }
                    resolve();
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

// Snooze the selected email
async function snoozeEmail(event: Office.AddinCommands.Event): Promise<void> {
  try {
    if (!snoozeFolderId) {
      await initializeSnoozeFolder();
    }
    
    const item = Office.context.mailbox.item;
    if (!item) {
      throw new Error('No email selected');
    }
    
    // Show taskpane for snooze options
    Office.context.ui.displayDialogAsync(
      `${window.location.origin}/taskpane.html?snooze=true`,
      { height: 40, width: 30 },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          showNotification('Error', 'Failed to open snooze options');
        }
        event.completed();
      }
    );
  } catch (error) {
    showNotification('Error', error instanceof Error ? error.message : 'Unknown error');
    event.completed();
  }
}

function showNotification(title: string, message: string): void {
  const details: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: `${title}: ${message}`,
    icon: "Icon.80x80",
    persistent: true,
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync(`snoozeNotification-${title}`, details);
}

// Register the function with Office.
Office.actions.associate("snoozeEmail", snoozeEmail);
