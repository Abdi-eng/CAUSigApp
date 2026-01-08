/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insert-sig-btn").onclick = runSignatureLogic;
  }
});

async function runSignatureLogic() {
  try {
    // 1. Get the Access Token from Outlook
    const token = await getAccessToken();
    
    // 2. Use the token to get Data from Microsoft Graph
    const userData = await getUserData(token);
    
    // 3. Generate the Signature with REAL data
    const signatureHtml = generateHtml(userData);

    // 4. Insert into Email
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureHtml,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById("status-message").innerText = "Success!";
        } else {
          document.getElementById("status-message").innerText = "Error applying signature.";
        }
      }
    );

  } catch (error) {
    console.error(error);
    // If SSO fails (e.g. user not logged in), show error
    document.getElementById("status-message").innerText = "Error: " + error.message;
  }
}

async function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessTokenAsync(
      { allowSignInPrompt: true, allowConsentPrompt: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          // If the error is code 13003, the user creates the token. 
          // But usually we just reject.
          reject(new Error("Could not get token: " + result.error.message));
        }
      }
    );
  });
}

async function getUserData(accessToken) {
  // Call Graph API
  const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mobilePhone,mail", {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    throw new Error("Graph API failed: " + response.statusText);
  }

  return await response.json();
}

function generateHtml(user) {
  // Use data from Graph, or fallback to empty string if missing
  const name = user.displayName;
  const title = user.jobTitle || "Staff Member";
  const phone = user.mobilePhone || "";
  const email = user.mail;

  return `
    <br>
    <table cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
        <tr>
            <td style="padding-right: 20px; border-right: 2px solid #0078d4; vertical-align: middle;">
                <img src="https://jolly-field-081c59603.2.azurestaticapps.net/assets/logo.png" width="80" height="80" style="display: block;">
            </td>
            <td style="padding-left: 20px; vertical-align: top;">
                <strong style="font-size: 18px; color: #2b579a;">${name}</strong><br>
                <span style="font-size: 13px; color: #666;">${title}</span>
                <br><br>
                <span style="color: #2b579a;">E:</span> <a href="mailto:${email}" style="text-decoration: none; color: #333;">${email}</a><br>
                ${phone ? `<span style="color: #2b579a;">P:</span> ${phone}<br>` : ""}
                <span style="color: #2b579a;">W:</span> <a href="https://www.cau.edu.tr" style="text-decoration: none; color: #333;">www.cau.edu.tr</a>
            </td>
        </tr>
    </table>
    <br>
  `;
}