
Office.onReady(() => {
  /** *
   * commands.js
   * This file contains the command functions for the Outlook2OneNote add-in.
   * It handles user interactions and communicates with the Office.js API.
   *  Note: This code assumes that you have already set up the necessary Office.js and Microsoft Graph API configurations.
   * 
   * Dependencies:
   * - Office.js
   * - Microsoft Graph API (for OneNote notebooks)
   * 
   * Global variables:
   * - info: Contains information about the Office host environment.
   * - document: The global document object for manipulating the DOM.
   * 
   * Usage:
   * - Call `action()` to show a notification when the add-in command is executed.  
   */

  console.log("Outlook2OneNote::commands::Office.onReady ")
if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
  }

});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  console.log("Outlook2OneNote::commands::action() ")
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
