/* global Office */

// ⬇️ PASTE YOUR AZURE FUNCTION URL HERE (Same one from taskpane.js) ⬇️
const API_URL = "https://api-cau-signature-d2edhphpf7hbg6g2.westeurope-01.azurewebsites.net/api/GetSignature"; 

Office.onReady();

// This function handles the "OnNewMessageCompose" event
function autoApplySignature(event) {
  setSignature(event);
}

async function setSignature(event) {
  try {
    const userEmail = Office.context.mailbox.userProfile.emailAddress;
    
    // Check if we already have the signature cached to be fast
    // (Optional optimization, but let's fetch fresh for now)
    
    // 1. Fetch from Azure Backend
    const separator = API_URL.includes("?") ? "&" : "?";
    const fullUrl = `${API_URL}${separator}email=${userEmail}`;
    
    const response = await fetch(fullUrl);
    
    if (!response.ok) {
        console.error("Backend Error");
        if (event) event.completed(); // Tell Outlook we are done even if failed
        return;
    }

    const signatureHtml = await response.text();

    // 2. Set the Signature
    // setSignatureAsync AUTOMATICALLY replaces/removes any existing Outlook signature
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureHtml,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        // 3. Signal completion to Outlook
        if (event) event.completed();
      }
    );

  } catch (error) {
    console.error(error);
    if (event) event.completed();
  }
}

// Register the function so the Manifest can find it
Office.actions.associate("autoApplySignature", autoApplySignature);