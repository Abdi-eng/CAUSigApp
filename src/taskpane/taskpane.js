/* global Office */

// ------------------------------------------------------------------
// ⬇️ PASTE YOUR AZURE FUNCTION URL HERE (Keep the ? at the end) ⬇️
// Example: "https://api-cau...net/api/GetSignature?code=abc..."
const API_URL = "https://api-cau-signature-d2edhphpf7hbg6g2.westeurope-01.azurewebsites.net/api/GetSignature"; 
// ------------------------------------------------------------------

let loadedHtml = ""; // Store the result globally

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insert-sig-btn").onclick = insertSignature;
    
    // Load data immediately when the pane opens
    loadSignatureFromBackend();
  }
});

async function loadSignatureFromBackend() {
    try {
        // 1. Get the current user's email from Outlook Client
        const userEmail = Office.context.mailbox.userProfile.emailAddress;
        
        // 2. Call your Azure Backend
        // We append &email=... to your Function URL
        // If your URL already has a '?', we use '&', otherwise '?'
        const separator = API_URL.includes("?") ? "&" : "?";
        const fullUrl = `${API_URL}${separator}email=${userEmail}`;

        const response = await fetch(fullUrl);

        if (!response.ok) {
            throw new Error("Backend Error: " + response.statusText);
        }

        // 3. Get the HTML from the response
        loadedHtml = await response.text();

        // 4. Show it in the preview box
        document.getElementById("signature-preview").innerHTML = loadedHtml;
        document.getElementById("status-message").innerText = ""; // Clear loading text

    } catch (error) {
        console.error(error);
        document.getElementById("signature-preview").innerHTML = `<p style="color:red">Error loading signature.</p>`;
        document.getElementById("status-message").innerText = error.message;
        document.getElementById("status-message").style.color = "red";
    }
}

function insertSignature() {
    if (!loadedHtml) {
        document.getElementById("status-message").innerText = "Please wait for signature to load...";
        return;
    }

    Office.context.mailbox.item.body.setSignatureAsync(
        loadedHtml,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("status-message").innerText = "Signature Inserted!";
                document.getElementById("status-message").style.color = "green";
                // Clear success message after 3 seconds
                setTimeout(() => document.getElementById("status-message").innerText = "", 3000);
            } else {
                document.getElementById("status-message").innerText = "Insert Failed: " + result.error.message;
                document.getElementById("status-message").style.color = "red";
            }
        }
    );
}