/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // 1. Load saved settings
    loadSettings();

    // 2. Add listeners to update preview instantly
    document.getElementById("user-job").addEventListener("input", updateSignature);
    document.getElementById("user-phone").addEventListener("input", updateSignature);

    // 3. Setup Insert Button
    document.getElementById("insert-sig-btn").onclick = insertSignature;
  }
});

function loadSettings() {
  const settings = Office.context.roamingSettings;
  const savedJob = settings.get("userJob");
  const savedPhone = settings.get("userPhone");

  if (savedJob) document.getElementById("user-job").value = savedJob;
  if (savedPhone) document.getElementById("user-phone").value = savedPhone;

  updateSignature();
}

function updateSignature() {
  // Get Inputs
  const job = document.getElementById("user-job").value;
  const phone = document.getElementById("user-phone").value;
  
  // Get Outlook Profile
  const userProfile = Office.context.mailbox.userProfile;
  const name = userProfile.displayName;
  const email = userProfile.emailAddress;

  // Save to Cloud
  const settings = Office.context.roamingSettings;
  settings.set("userJob", job);
  settings.set("userPhone", phone);
  settings.saveAsync(); 

  // Generate HTML
  const html = generateHtml(name, email, job, phone);

  // Show Preview
  document.getElementById("signature-preview").innerHTML = html;
}

function insertSignature() {
  const job = document.getElementById("user-job").value;
  const phone = document.getElementById("user-phone").value;
  const userProfile = Office.context.mailbox.userProfile;
  
  const html = generateHtml(userProfile.displayName, userProfile.emailAddress, job, phone);

  Office.context.mailbox.item.body.setSignatureAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("status-message").innerText = "Signature Inserted!";
        setTimeout(() => document.getElementById("status-message").innerText = "", 3000);
      }
    }
  );
}

/*function generateHtml(name, email, job, phone) {
  // 1. Force Name to Uppercase
  const displayName = name.toUpperCase();
  
  // 2. Defaults
  const displayJob = job || "Unvan Giriniz"; // "Enter Title" in Turkish
  const displayPhone = phone || "";

  // 3. The Design (Matches your Screenshot)
  // We use your Azure URL for the image
  return `
    <br>
    <table cellpadding="0" cellspacing="0" style="font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #333333; text-align: left;">
        <tr>
            <!-- LEFT SIDE: LOGO -->
            <td style="padding-right: 20px; vertical-align: top;">
                <img src="https://jolly-field-081c59603.2.azurestaticapps.net/assets/logo.png" width="155" height="155" style="display: block;">
            </td>
            
            <!-- RIGHT SIDE: TEXT -->
            <td style="vertical-align: top; line-height: 1.4;">
                <!-- NAME (Uppercase & Bold) -->
                <strong style="font-size: 15px; text-transform: uppercase; color: #000;">${displayName}</strong>
                <br>
                
                <!-- JOB TITLE (Bold) -->
                <strong style="font-size: 14px; color: #000;">${displayJob}</strong>
                <br>
                
                <!-- PHONE -->
                <span style="font-size: 13px;">Tel: ${displayPhone}</span>
                <br>
                
                <!-- EMAIL (Blue Link) -->
                <a href="mailto:${email}" style="font-size: 13px; color: #0563C1; text-decoration: underline;">${email}</a>
                <br>
                
                <!-- ADDRESS (Hardcoded) -->
                <span style="font-size: 13px;">Fazıl Küçük Cad. No. 80</span>
                <br>
                <span style="font-size: 13px;">Ozanköy, Girne – Kuzey Kıbrıs</span>
                <br>
                
                <!-- WEBSITE -->
                <a href="https://www.cau.edu.tr" style="font-size: 13px; color: #0563C1; text-decoration: underline;">www.cau.edu.tr</a>
            </td>
        </tr>
    </table>
    <br>
  `;
}*/
function generateHtml(name, email, job, phone) {
  const displayName = name.toUpperCase();
  
  // Logic to hide empty lines
  const jobLine = job ? `<strong style="font-size: 14px; color: #000;">${job}</strong><br>` : "";
  const phoneLine = phone ? `<span style="font-size: 13px;">Tel: ${phone}</span><br>` : "";

  return `
    <br>
    <!-- MAIN TABLE: Fixed width of 600px to prevent wrapping -->
    <table width="600" cellpadding="0" cellspacing="0" border="0" style="width: 600px; font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #333333; text-align: left;">
        <tr>
            <!-- LEFT SIDE: LOGO (Fixed width 155px) -->
            <!-- The 'border-right' adds the BLUE LINE -->
            <td width="130" style="width: 130px; padding-right: 15px; border-right: 2px solid #0563C1; vertical-align: top;">
                <!-- Fixed Image Dimensions -->
                <img src="https://jolly-field-081c59603.2.azurestaticapps.net/assets/logo.png" width="155" height="155" style="display: block; width: 155px; height: 155px;">
            </td>
            
            <!-- RIGHT SIDE: TEXT -->
            <td style="padding-left: 15px; vertical-align: top; line-height: 1.4;">
                <strong style="font-size: 15px; text-transform: uppercase; color: #000;">${displayName}</strong>
                <br>
                
                ${jobLine}
                
                ${phoneLine}
                
                <a href="mailto:${email}" style="font-size: 13px; color: #0563C1; text-decoration: underline;">${email}</a>
                <br>
                
                <span style="font-size: 13px;">Fazıl Küçük Cad. No. 80</span>
                <br>
                <span style="font-size: 13px;">Ozanköy, Girne – Kuzey Kıbrıs</span>
                <br>
                
                <a href="https://www.cau.edu.tr" style="font-size: 13px; color: #0563C1; text-decoration: underline;">www.cau.edu.tr</a>
            </td>
        </tr>
    </table>
    <br>
  `;
}