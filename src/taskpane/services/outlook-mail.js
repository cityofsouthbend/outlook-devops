/* Office.js mail interaction helpers */

export function getMode() {
  const item = Office.context.mailbox.item;
  if (!item) return "unknown";
  return typeof item.subject === "object" ? "compose" : "read";
}

export function getReadMessageInfo() {
  const item = Office.context.mailbox.item;
  return {
    subject: item.subject,
    from: item.from,
    attachments: item.attachments || [],
    itemId: item.itemId,
  };
}

export function getCurrentUserName() {
  return Office.context.mailbox.userProfile.displayName;
}

export function getEmailBody() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

export function getAttachmentContent(attachmentId, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (!settled) {
        settled = true;
        reject(new Error("getAttachmentContentAsync timed out"));
      }
    }, timeoutMs);

    Office.context.mailbox.item.getAttachmentContentAsync(
      attachmentId,
      (result) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

export function getRestToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

export function getRestItemId() {
  return Office.context.mailbox.convertToRestId(
    Office.context.mailbox.item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );
}

export async function getAttachmentContentViaRest(attachmentId) {
  const token = await getRestToken();
  const restUrl = Office.context.mailbox.restUrl;
  const restItemId = getRestItemId();
  const restAttachmentId = Office.context.mailbox.convertToRestId(
    attachmentId,
    Office.MailboxEnums.RestVersion.v2_0
  );

  const url = `${restUrl}/v2.0/me/messages/${restItemId}/attachments/${restAttachmentId}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    throw new Error(`REST attachment fetch failed: ${res.status} ${res.statusText}`);
  }

  const data = await res.json();
  return {
    content: data.ContentBytes,
    format: Office.MailboxEnums.AttachmentContentFormat.Base64,
    contentType: data.ContentType,
  };
}

export function insertHtmlIntoCompose(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(asyncResult.error.message));
        } else {
          resolve();
        }
      }
    );
  });
}
