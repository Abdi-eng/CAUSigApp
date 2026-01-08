/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insert-sig-btn").onclick = runSignatureLogic;
  }
});

function runSignatureLogic() {
  // 1. Get Basic Data (No SSO required)
  const userProfile = Office.context.mailbox.userProfile;
  const displayName = userProfile.displayName;
  const email = userProfile.emailAddress;
  
  // 2. Fallback for Title/Phone (Since we can't reach Exchange without SSO)
  // You can add input boxes for these in HTML later if you want
  const jobTitle = "Staff Member"; 
  const phone = ""; 

  // 3. Generate HTML
  const signatureHtml = `
    <br>
    <table cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
        <tr>
            <td style="padding-right: 20px; border-right: 2px solid #0078d4; vertical-align: middle;">
                <img src="https://jolly-field-081c59603.2.azurestaticapps.net/assets/logo.png" width="80" height="80">
            </td>
            <td style="padding-left: 20px;">
                <strong style="font-size: 18px; color: #2b579a;">${displayName}</strong><br>
                <span>${jobTitle}</span><br><br>
                <a href="mailto:${email}">${email}</a>
            </td>
        </tr>
    </table>
  `;

  // 4. Insert
  Office.context.mailbox.item.body.setSignatureAsync(
    signatureHtml,
    { coercionType: Office.CoercionType.Html },
    (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("status-message").innerText = "Success!";
        }
    }
  );
}