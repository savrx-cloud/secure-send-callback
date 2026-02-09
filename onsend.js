
// src/onsend.js
Office.onReady(() => {});

async function enforceSecureOnSend(evt) {
  try {
    const item = Office.context.mailbox.item;

    const wantSecure = await new Promise((resolve, reject) => {
      item.loadCustomPropertiesAsync((cpRes) => {
        if (cpRes.status !== Office.AsyncResultStatus.Succeeded) return reject(cpRes.error);
        const props = cpRes.value;
        resolve(String(props.get("sendSecure") || "").toLowerCase() == "true");
      });
    });

    let needsFix = false;
    let newSubject = "";

    await new Promise((resolve, reject) => {
      item.subject.getAsync({}, (getRes) => {
        if (getRes.status !== Office.AsyncResultStatus.Succeeded) return reject(getRes.error);
        const subj = getRes.value || "";
        const hasPrefix = /^secure:\s?/i.test(subj);
        if (wantSecure && !hasPrefix) {
          needsFix = true;
          newSubject = `Secure: ${subj}`;
        } else if (hasPrefix) {
          needsFix = /^secure:\s?/i.test(subj) && !/^Secure:\s/.test(subj);
          newSubject = subj.replace(/^secure:\s?/i, "Secure: ");
        }
        resolve();
      });
    });

    if (needsFix) {
      await new Promise((resolve, reject) => {
        item.subject.setAsync(newSubject, (setRes) => {
          if (setRes.status !== Office.AsyncResultStatus.Succeeded) return reject(setRes.error);
          resolve();
        });
      });
    }

    evt.completed({ allowEvent: true });
  } catch (e) {
    Office.context.mailbox.item.notificationMessages.addAsync("sendsecure-block", {
      type: "errorMessage",
      message: `Send Secure blocked: ${e.message || e}. Please try again.`
    });
    evt.completed({ allowEvent: false });
  }
}

Office.actions.associate("enforceSecureOnSend", enforceSecureOnSend);
