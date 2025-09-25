/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    showPhishingWarning();
  }
});

function showPhishingWarning() {
  const item = Office.context.mailbox.item;
  
  if (item) {
    item.notificationMessages.replaceAsync("phishWarning", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "⚠ Đây là email đáng ngờ (phishing test)!",
      icon: "icon-16", 
      persistent: true
    });
  }
}
