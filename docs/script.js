let lastKnownBody = "";

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    startMonitoringForCapitalization();
  }
});

function startMonitoringForCapitalization() {
  setInterval(() => {
    Office.context.mailbox.item.body.getAsync("html", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const currentBody = result.value;
        if (currentBody !== lastKnownBody && /sprchrgr/.test(currentBody)) {
          const updatedBody = currentBody.replace(/sprchrgr/g, "SPRCHRGR");
          Office.context.mailbox.item.body.setAsync(updatedBody, { coercionType: "html" }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              lastKnownBody = updatedBody;
              console.log("Auto-capitalized SPRCHRGR.");
            }
          });
        } else {
          lastKnownBody = currentBody;
        }
      }
    });
  }, 2000);
}
