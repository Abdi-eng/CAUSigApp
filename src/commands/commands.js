/* global Office */

const API_URL = "https://api-cau-signature-d2edhphpf7hbg6g2.westeurope-01.azurewebsites.net/api/GetSignature"; 

Office.onReady(() => {
    // Office is ready
});

/**
 * This function is the entry point for the Launch Event.
 * It must be fast.
 */
async function autoApplySignature(event) {
    console.log("Auto-apply signature event triggered.");
    
    try {
        const userEmail = Office.context.mailbox.userProfile.emailAddress;
        
        // 1. Fetch from Azure Backend with a timeout safety
        const separator = API_URL.includes("?") ? "&" : "?";
        const fullUrl = `${API_URL}${separator}email=${encodeURIComponent(userEmail)}`;
        
        // We use an AbortController to ensure we don't hang and hit the 5s limit
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 4000); // 4 second timeout

        const response = await fetch(fullUrl, { signal: controller.signal });
        clearTimeout(timeoutId);

        if (!response.ok) {
            throw new Error("Backend responded with error");
        }

        const signatureHtml = await response.text();

        // 2. Set the Signature
        Office.context.mailbox.item.body.setSignatureAsync(
            signatureHtml,
            { coercionType: Office.CoercionType.Html },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("setSignatureAsync failed: " + asyncResult.error.message);
                } else {
                    console.log("Signature applied successfully.");
                }
                // 3. Signal completion to Outlook (MANDATORY)
                event.completed();
            }
        );

    } catch (error) {
        console.error("Signature App Error: ", error);
        // Even if it fails, we MUST call event.completed() to release the process
        event.completed();
    }
}
// IMPORTANT: This association must happen at the global level
Office.actions.associate("autoApplySignature", autoApplySignature);