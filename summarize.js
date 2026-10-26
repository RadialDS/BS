// Pure logic (no Excel calls) so it can be tested outside Excel.
function summarize(values) {
  const hdr = values[0].map(h => String(h).trim().toLowerCase());
  const ix = n => hdr.indexOf(n);
  const c = { loc: ix("location_name"), st: ix("appt_status"), bb: ix("buy_and_bill"),
              ph: ix("pharmacy"), dose: ix("dose"), mrn: ix("med_rec_nbr") };
  const missing = Object.entries(c).filter(([k, v]) => v < 0).map(([k]) => k);
  if (missing.length) throw new Error("Missing column(s): " + missing.join(", "));

  const clinics = {};   // clinic -> {d56, d84, other:{dose:count}}
  const mismatches = [];
  for (const r of values.slice(1)) {
    const loc = String(r[c.loc]).trim();
    if (!loc) continue;
    clinics[loc] = clinics[loc] || { d56: 0, d84: 0, other: {} };
    if (String(r[c.st]).trim().toLowerCase() !== "kept") continue;
    const flagBB = Number(r[c.bb]) === 1;
    const isBesse = String(r[c.ph]).trim().toLowerCase() === "besse";
    if (!flagBB && !isBesse) continue;
    if (flagBB !== isBesse)
      mismatches.push([r[c.mrn], loc, String(r[c.ph]).trim(), flagBB ? 1 : 0]);
    const mg = parseFloat(String(r[c.dose]).replace(/[^0-9.]/g, ""));
    if (mg === 56) clinics[loc].d56++;
    else if (mg === 84) clinics[loc].d84++;
    else {
      const label = String(r[c.dose]).trim() || "blank";
      clinics[loc].other[label] = (clinics[loc].other[label] || 0) + 1;
    }
  }
  const rows = Object.keys(clinics).sort((a, b) => a.localeCompare(b)).map(k => {
    const o = Object.entries(clinics[k].other).map(([d, n]) => `${d} x${n}`).join(", ");
    return [k, clinics[k].d56, clinics[k].d84, o ? "CHECK: " + o : ""];
  });
  return { rows, mismatches };
}
if (typeof module !== "undefined") module.exports = { summarize };
