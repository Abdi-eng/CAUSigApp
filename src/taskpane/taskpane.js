/* global Office */

let userSignature = ""; // Store it globally

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // 1. Generate the signature immediately
    generateSignature();
    
    // 2. Attach the "Insert" button to the insert function
    document.getElementById("insert-sig-btn").onclick = insertSignature;
  }
});

function generateSignature() {
  const userProfile = Office.context.mailbox.userProfile;
  const userName = userProfile.displayName; 
  const userEmail = userProfile.emailAddress;

  // Define the HTML
  userSignature = `
    <table style="font-family: Arial; color: #333;">
        <tr>
            <td style="border-right: 2px solid #0078d4; padding-right: 10px;">
                <!-- Use your LOCALHOST URL for now -->
                <img src="https://localhost:3000/assets/logo.png" width="60">
            </td>
            <td style="padding-left: 10px;">
                <strong>${userName}</strong><br>
                <a href="mailto:${userEmail}" style="text-decoration:none; color:#0078d4">${userEmail}</a><br>
                <span>Cyprus Aydin University</span>
            </td>
        </tr>
    </table>
  `;

  // Show it in the taskpane preview box
  document.getElementById("signature-preview").innerHTML = userSignature;
}

function insertSignature() {
  // Insert the stored HTML
  Office.context.mailbox.item.body.setSignatureAsync(
    userSignature,
    { coercionType: Office.CoercionType.Html },
    (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("status-message").innerText = "Signature Inserted!";
        }
    }
  );
}