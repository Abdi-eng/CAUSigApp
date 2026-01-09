/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insert-sig-btn").onclick = runSignatureLogic;
  }
});

async function runSignatureLogic() {
  try {
    // 1. Get Basic Data (Always works)
    const userProfile = Office.context.mailbox.userProfile;
    let name = userProfile.displayName;
    let email = userProfile.emailAddress;
    let jobTitle = "Staff Member"; // Default
    let phone = ""; 

    // 2. Generate Signature
    const signatureHtml = generateHtml(name, email, jobTitle, phone);

    // 3. Insert
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureHtml,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById("status-message").innerText = "Success!";
        }
      }
    );

  } catch (error) {
    console.error(error);
    document.getElementById("status-message").innerText = "Error: " + error.message;
  }
}

function generateHtml(name, email, title, phone) {
  return `
    <br>
    <table cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
        <tr>
            <td style="padding-right: 20px; border-right: 2px solid #0078d4; vertical-align: middle;">
                <img src="https://jolly-field-081c59603.2.azurestaticapps.net/assets/logo.png" width="80" height="80" style="display: block;">
            </td>
            <td style="padding-left: 20px;">
                <strong style="font-size: 18px; color: #2b579a;">${name}</strong><br>
                <span>${title}</span><br><br>
                <a href="mailto:${email}" style="text-decoration:none; color:#333;">${email}</a>
                ${phone ? `<br><span>P: ${phone}</span>` : ""}
            </td>
        </tr>
    </table>
  `;
}