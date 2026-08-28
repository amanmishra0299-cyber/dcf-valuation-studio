/**
 * ui_test.js -- renders the real page in jsdom, drives it with the real API,
 * and asserts on the resulting DOM. This exercises the shipped JavaScript,
 * not a reimplementation of it.
 *
 * Requires the Flask server on 127.0.0.1:8000.
 * Run: node ui_test.js
 */
const fs = require("fs");
const { JSDOM, VirtualConsole } = require("jsdom");

const HTML = fs.readFileSync(__dirname + "/index.html", "utf8");
const BASE = "http://127.0.0.1:8080";

let pass = 0, fail = 0;
const check = (name, cond, detail) => {
  if (cond) { pass++; console.log("  ok   " + name); }
  else { fail++; console.log("  FAIL " + name + (detail ? "   " + detail : "")); }
};

function makeDom() {
  const errors = [];
  const vc = new VirtualConsole();
  vc.on("jsdomError", e => errors.push("jsdomError: " + (e.stack || e.message)));
  vc.on("error", (...a) => errors.push("console.error: " + a.join(" ")));

  const dom = new JSDOM(HTML, {
    runScripts: "dangerously",
    pretendToBeVisual: true,
    url: BASE + "/",
    virtualConsole: vc,
    resources: undefined,
  });

  // real network, relative to the Flask app
  dom.window.fetch = (url, opts) =>
    fetch(url.toString().startsWith("http") ? url.toString() : BASE + url, opts);

  dom.window.URL.createObjectURL = () => "blob:stub";
  dom.window.URL.revokeObjectURL = () => {};
  dom.window.alert = m => { throw new Error("alert(): " + m); };
  return { dom, errors };
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// visible text only -- the inline <script> legitimately contains "isNaN"
const visibleText = d => [...d.querySelectorAll("body *:not(script):not(style)")]
  .filter(el => el.children.length === 0)
  .map(el => el.textContent).join(" ");

async function analyze(ticker) {
  const { dom, errors } = makeDom();
  const w = dom.window, d = w.document;
  d.getElementById("ticker").value = ticker;
  d.getElementById("go").click();

  // wait for the result template to be mounted
  for (let i = 0; i < 200 && !d.getElementById("vps"); i++) await sleep(150);
  if (!d.getElementById("vps")) throw new Error("timed out rendering " + ticker);
  await sleep(300); // let the remaining panels paint
  return { dom, errors, d, w };
}

(async () => {
  console.log("\n[A] RELIANCE (non-financial)");
  let { dom, errors, d } = await analyze("RELIANCE");
  const txt = s => (d.getElementById(s)?.textContent || "").trim();

  check("no JS errors on load+render", errors.length === 0, errors.slice(0, 3).join(" | "));
  check("company name rendered", /Reliance/i.test(txt("cname")), txt("cname"));
  check("intrinsic value rendered", /^₹[\d,\.]+$/.test(txt("vps")), txt("vps"));
  check("market price rendered", /^₹[\d,\.]+$/.test(txt("cprice")), txt("cprice"));
  check("verdict pill rendered", /valued/i.test(txt("vPill")), txt("vPill"));
  check("6 KPI tiles", d.querySelectorAll("#kpis .kpi").length === 6,
        d.querySelectorAll("#kpis .kpi").length);
  check("WACC table built", /WACC/.test(txt("waccBox")), txt("waccBox").slice(0, 60));
  check("equity bridge built", /Enterprise value/.test(txt("bridge")));
  const projRows = d.querySelectorAll("#proj tr").length;
  check("projection table has header + 10 years + terminal", projRows === 12, projRows);
  check("projection has real numbers", !/NaN/.test(d.getElementById("proj").textContent));
  check("sensitivity grid = header + 7 WACC rows", d.querySelectorAll("#heat tr").length === 8,
        d.querySelectorAll("#heat tr").length);
  check("sensitivity cells are numbers",
        !/—|NaN/.test([...d.querySelectorAll("#heat td")].slice(0, 5).map(t => t.textContent).join("")),
        d.querySelector("#heat").textContent.slice(0, 60));
  check("scenarios table built", /Bull/.test(txt("scenBox")), txt("scenBox").slice(0, 40));
  check("implied expectation rendered", /%/.test(txt("impliedBox")), txt("impliedBox").slice(0, 50));
  check("DDM panel built", /Cost of equity/.test(txt("ddmBox")));
  check("comps table built", d.querySelectorAll("#peersT tr").length > 3,
        d.querySelectorAll("#peersT tr").length);
  check("peer multiples implied", /EV\/EBITDA|P\/E/.test(txt("relImp")), txt("relImp").slice(0, 40));
  check("statements rendered", d.querySelectorAll("#stmts table").length >= 3,
        d.querySelectorAll("#stmts table").length);
  check("provenance table built", d.querySelectorAll("#provT tr").length > 10,
        d.querySelectorAll("#provT tr").length);
  check("football field SVG drawn", d.querySelectorAll("#ff svg line, #ff svg rect").length > 3,
        d.querySelectorAll("#ff svg *").length);
  check("projection chart SVG drawn", d.querySelectorAll("#projChart svg polyline").length === 2);
  const vt = visibleText(d);
  check("no NaN in the rendered output", !/NaN/.test(vt),
        (vt.match(/.{0,60}NaN.{0,30}/) || [""])[0]);
  check("no 'undefined' leaking into the output", !/undefined/.test(vt),
        (vt.match(/.{0,60}undefined.{0,30}/) || [""])[0]);
  check("assumption form populated", d.querySelectorAll("[data-k]").length >= 20,
        d.querySelectorAll("[data-k]").length);
  const betaInput = d.querySelector('[data-k="beta"]');
  check("beta prefilled from data", betaInput && betaInput.value !== "" && betaInput.value !== "1",
        betaInput && betaInput.value);

  // changing an assumption and recalculating must move the answer
  const before = txt("vps");
  d.querySelector('[data-k="terminal_growth_pct"]').value = "9";
  const w = dom.window;
  await w.eval("recalc()");
  for (let i = 0; i < 60 && txt("vps") === before; i++) await sleep(150);
  check("recalc changes the valuation", txt("vps") !== before, `${before} -> ${txt("vps")}`);
  check("still no JS errors after recalc", errors.length === 0, errors.slice(0, 3).join(" | "));
  dom.window.close();

  console.log("\n[B] HDFCBANK (financial -> residual income)");
  ({ dom, errors, d } = await analyze("HDFCBANK"));
  check("no JS errors", errors.length === 0, errors.slice(0, 3).join(" | "));
  check("bank/NBFC warning shown", /bank \/ NBFC/.test(d.getElementById("warnBox").textContent),
        d.getElementById("warnBox").textContent.slice(0, 80));
  check("headline uses residual income",
        /Residual income model/.test(d.getElementById("vpsSub").textContent),
        d.getElementById("vpsSub").textContent.slice(0, 70));
  check("residual income card visible",
        d.getElementById("rimCard").style.display === "", d.getElementById("rimCard").style.display);
  check("RIM table built", d.querySelectorAll("#rimT tr").length > 5,
        d.querySelectorAll("#rimT tr").length);
  check("RIM summary built", /Implied P\/B/.test(d.getElementById("rimSum").textContent));
  check("bank comps are banks", /ICICIBANK|AXISBANK|KOTAKBANK/.test(d.getElementById("peersT").textContent),
        d.getElementById("peersT").textContent.slice(0, 80));
  check("no NaN in bank view", !/NaN/.test(visibleText(d)));
  dom.window.close();

  console.log("\n[C] bad ticker");
  {
    const { dom } = makeDom();
    const d = dom.window.document;
    d.getElementById("ticker").value = "NOTAREALCO";
    d.getElementById("go").click();
    for (let i = 0; i < 120 && !d.querySelector(".err"); i++) await sleep(150);
    const err = d.querySelector(".err");
    check("error message shown, not a blank page", !!err, d.body.textContent.slice(0, 80));
    if (err) check("error text is helpful", /screener\.in|symbol/i.test(err.textContent),
                   err.textContent.slice(0, 100));
    dom.window.close();
  }

  console.log(`\n${"=".repeat(58)}\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
