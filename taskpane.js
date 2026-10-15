Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;
  document.getElementById("run").onclick = run;
});

const out = (msg) => (document.getElementById("out").textContent = msg);

async function run() {
  const prompt = document.getElementById("prompt").value.trim();
  if (!prompt) return out("Write a request first.");

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRange();
    used.load("address, values");
    await context.sync();

    // TODO: send `prompt` and `used.values` to your own backend, which calls the model.
    // The backend returns a suggested formula and a target cell. For now this is a demo.
    out(`Sheet: ${used.address}\nRows read: ${used.values.length}\n(Connect your backend to get real answers.)`);
  });
}
