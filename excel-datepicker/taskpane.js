Office.onReady(() => { });

async function insertDateTime() {
  const d = document.getElementById("date").value;
  const t = document.getElementById("time").value || "00:00";

  if (!d) return;

  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.values = [[`${d} ${t}`]];
    cell.numberFormatLocal = [["dd.mm.yyyy hh:mm"]];
    await context.sync();
  });
}

// вызов кнопки на ленте
function openPicker() {
  Office.context.ui.displayDialogAsync(
    window.location.origin + "/taskpane.html",
    { height: 50, width: 25 }
  );
}

if (typeof module !== "undefined") {
  module.exports = { openPicker };
}
