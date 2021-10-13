/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function insertTemplate(event: Office.AddinCommands.Event) {
  const htmlTemlpate = `<div><img src="https://loremflickr.com/cache/resized/65535_51377258275_023a2ab94b_320_240_nofilter.jpg"></img><div>`;
  const options = { coercionType: 'html' }

  Office.context.mailbox.item.body.setAsync(htmlTemlpate, options, (asyncResult) => {
    console.log('setAsync performed', asyncResult)
  })
}


function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;

g.insertTemplate = insertTemplate
