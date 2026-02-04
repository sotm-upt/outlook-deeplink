Office.onReady(() => {
  // If needed, perform initialization here
});

// The function name must match <FunctionName> in the manifest
function generateDeepLink(event) {
  // 1. Get the Item ID of the current message
  const item = Office.context.mailbox.item;
  
  // Note: We use the EWS ID (itemId) directly. The Flow handles encoding.
  const emailId = item.itemId;

  // 2. Prepare the payload for Power Automate
  const payload = {
    emailId: emailId
  };

  // 3. Call the Power Automate HTTP Trigger
  // REPLACE 'YOUR_POWER_AUTOMATE_URL' with the URL from Part 1
  const flowUrl = "https://defaultbf3289d6ee434e0d9204110b7e9003.dc.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/4b08bb7bc1b946679e67160a0e7b2e70/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NEia-ntvdX8Mu4OnKwG33c9aarEB-DHD3vOn9nwsOO8";

  fetch(flowUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  })
  .then((response) => response.json())
  .then((data) => {
    // 4. Display the generated Deep Link
    // We use a dialog or notification to show the result
    const message = "Deep Link Generated: " + data.deepLink;
    
    // Copy to clipboard or show to user
    // For this example, we display a notification
    Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Link generated! Check console or implement clipboard copy.",
      icon: "Icon.80x80",
      persistent: false
    });
    
    console.log(data.deepLink); // In a real add-in, you might open a dialog to let the user copy this
  })
  .catch((error) => {
    console.error("Error calling flow:", error);
    Office.context.mailbox.item.notificationMessages.replaceAsync("notify", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Failed to generate link.",
    });
  })
  .finally(() => {
    // Always call event.completed() to stop the button's loading state
    event.completed();
  });
}