// taskpane.js
(async () => {
  // helper: chuẩn hoá địa chỉ email (từ "Name <email@domain>" hoặc plain)
  function extractEmail(raw) {
    if (!raw) return "";
    // nếu có <> thì lấy trong ngoặc nhọn
    const m = raw.match(/<([^>]+)>/);
    const email = m ? m[1] : raw;
    return email.trim().toLowerCase();
  }

  async function loadLists() {
    try {
      const res = await fetch("/email-lists.json", { cache: "no-store" });
      if (!res.ok) throw new Error("Không load được email-lists.json: " + res.status);
      return await res.json();
    } catch (err) {
      console.error("Lỗi load email-lists:", err);
      return { whitelist: [], blacklist: [] };
    }
  }

  function checkEmailAgainstLists(email, lists) {
    if (!email) return "NEUTRAL";
    const e = email.toLowerCase();
    if (Array.isArray(lists.whitelist) && lists.whitelist.map(x=>x.toLowerCase()).includes(e)) return "SAFE";
    if (Array.isArray(lists.blacklist) && lists.blacklist.map(x=>x.toLowerCase()).includes(e)) return "PHISHING WARNING";
    return "NEUTRAL";
  }

  Office.onReady(async () => {
    console.log("Office.onReady fired ✅");
    const lists = await loadLists();
    console.log("Loaded lists:", lists);

    const item = Office.context.mailbox.item;
    console.log("Item object: ", item);

    // Lấy sender (một số host trả item.from.emailAddress, một số trả item.from.raw)
    let senderRaw = (item.from && (item.from.emailAddress || item.from.displayName)) || item.from || "";
    // nếu item.from là object khác, thử stringify cho debug
    if (typeof senderRaw === "object") {
      console.log("item.from object:", item.from);
      senderRaw = item.from.emailAddress || item.from.displayName || JSON.stringify(item.from);
    }

    const senderEmail = extractEmail(senderRaw);
    const subject = item.subject || "(Không có tiêu đề)";

    // Hiển thị nhanh lên UI
    document.getElementById("sender").innerText = senderEmail || "Không xác định";
    document.getElementById("subject").innerText = subject;

    // Lấy body preview (async)
    item.body.getAsync("text", (result) => {
      console.log("Body result:", result);
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const preview = (result.value || "").substring(0, 500);
        document.getElementById("bodyPreview").innerText = preview;
      } else {
        document.getElementById("bodyPreview").innerText = "Không lấy được nội dung.";
      }
    });

    // Chạy check sau khi đã có senderEmail và lists
    const checkResult = checkEmailAgainstLists(senderEmail, lists);
    document.getElementById("result").innerText = checkResult;

    // Hiển thị notification trên email
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync("phishCheck", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: checkResult === "PHISHING WARNING" ? "⚠️ Email khả nghi" : (checkResult === "SAFE" ? "✅ Email an toàn" : "ℹ️ Trung tính"),
        icon: "icon16",
        persistent: false
      }, function(asyncResult) {
        console.log("replaceAsync result:", asyncResult);
      });
    } catch (err) {
      console.error("Không thể hiển thị notification:", err);
    }

    // Optionally auto open taskpane (if code running from event-based function, else not needed here)
    // Office.addin.showAsTaskpane(); // chỉ gọi nếu bạn muốn bật taskpane từ commands
  });
})();
