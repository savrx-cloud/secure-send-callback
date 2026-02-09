
// src/commands.js
Office.onReady(() => {});

async function markSecureAndPrefix(evt) {
  try {
    const item = Office.context.mailbox.item;

    await new Promise((resolve, reject) => {
      item.subject.getAsync({}, (getRes) => {
        if (getRes.status !== Office.AsyncResultStatus.Succeeded) return reject(getRes.error);
        let subj = getRes.value || "";
        if (!/^secure:\s?/i.test(subj)) {
          subj = `Secure: ${subj}`;
          item.subject.setAsync(subj, (setRes) => {
            if (setRes.status !== Office.AsyncResultStatus.Succeeded) return reject(setRes.error);
            resolve();
          });
        } else {
          // Normalize to "Secure: "
          subj = subj.replace(/^secure:\s?/i, "Secure: ");
          item.subject.setAsync(subj, (setRes) => {
            if (setRes.status !== Office.AsyncResultStatus.Succeeded) return reject(setRes.error);
            resolve();
          });
        }
      });
    });

    await new Promise((resolve, reject) => {
      item.loadCustomPropertiesAsync((cpRes) => {
        if (cpRes.status !== Office.AsyncResultStatus.Succeeded) return reject(cpRes.error);
        const props = cpRes.value;
        props.set("sendSecure", "true");
        props.saveAsync((saveRes) => {
          if (saveRes.status !== Office.AsyncResultStatus.Succeeded) return reject(saveRes.error);
          resolve();
        });
      });
    });

    item.notificationMessages.replaceAsync("sendsecure",
      { type: "informationalMessage", message: "Secure subject set. Click Send to deliver encrypted.", icon: "icon-16", persistent: false }
    );
  } catch (e) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("sendsecureerr",
      { type: "errorMessage", message: `Send Secure: ${e.message || e}` }
    );
  } finally {
    evt?.completed?.();
  }
}

Office.actions.associate("markSecureAndPrefix", markSecureAndPrefix);
