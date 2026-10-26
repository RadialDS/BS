async function run() {
  const msg = document.getElementById("msg");
  msg.textContent = "Working...";
  try {
    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const src = sheets.items.find(s => /^buyandbillschedule_/i.test(s.name));
      if (!src) throw new Error('No sheet named like "buyandbillschedule_YYYYMMDD" found.');
      const used = src.getUsedRange(true);
      used.load("values");
      await ctx.sync();

      const { rows, mismatches } = summarize(used.values);
      const date = src.name.split("_")[1] || "";
      const name = "BB Summary " + date;
      const old = sheets.getItemOrNullObject(name);
      await ctx.sync();
      if (!old.isNullObject) old.delete();
      const ws = sheets.add(name);

      const n = rows.length;
      ws.getRange("A1:D1").values = [["Clinic", "56 mg", "84 mg", "Flag: other dose"]];
      ws.getRange("A1:D1").format.font.bold = true;
      ws.getRange("A1:D1").format.fill.color = "#D9E1F2";
      ws.getRange(`A2:D${n + 1}`).values = rows;
      ws.getRange(`B2:C${n + 1}`).numberFormat = [["0"]];
      ws.getRange(`B1:C${n + 1}`).format.horizontalAlignment = "Center";
      const flagRows = ws.getRange(`D2:D${n + 1}`);
      const cf = flagRows.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula = '=LEN($D2)>0';
      cf.custom.format.fill.color = "#FFC7CE";
      cf.custom.format.font.color = "#9C0006";

      if (mismatches.length) {
        ws.getRange("F1:I1").values = [["Mismatch: MRN", "Clinic", "Pharmacy", "buy_and_bill"]];
        ws.getRange("F1:I1").format.font.bold = true;
        ws.getRange("F1:I1").format.fill.color = "#FFEB9C";
        ws.getRange(`F2:I${mismatches.length + 1}`).values = mismatches;
      }
      ws.getRange("A:I").format.autofitColumns();
      ws.activate();
      await ctx.sync();
      msg.textContent = `Created "${name}": ${n} clinics, ${mismatches.length} mismatch(es), ` +
        `${rows.filter(r => r[3]).length} dose flag(s).`;
    });
  } catch (e) { msg.textContent = "Error: " + e.message; }
}
Office.onReady(() => { document.getElementById("go").onclick = run; });
