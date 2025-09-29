// commands.js
Office.onReady(() => console.log("Commands ready"));

function onMessageRead(event) {
  console.log("onMessageRead fired");
  try {
    // show taskpane (nếu muốn)
    Office.addin.showAsTaskpane();
  } catch (err) {
    console.warn("Không thể mở taskpane tự động:", err);
  }
  event.completed();
}

// export (nếu cần bundler)
if (typeof module !== "undefined") module.exports = { onMessageRead };
