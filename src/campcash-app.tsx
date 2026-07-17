import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  Wallet, Bell, MoreVertical, Home, BarChart2, PieChart as PieChartIcon, Settings,
  ArrowDownCircle, ArrowUpCircle, Search, SlidersHorizontal, ChevronDown, ChevronUp,
  Plus, Minus, Calendar, Moon, Sun, Trash2, User, Download, Upload, FileText, FileSpreadsheet,
  AlertTriangle, Sparkles, Landmark, Copy, Cloud
} from "lucide-react";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from "recharts";
import * as XLSX from "xlsx";

const DEFAULT_SOURCES = [
  "Parents / Family",
  "Scholarship",
  "Stipend / Internship Income",
  "Refund",
  "FD / Savings Withdrawal",
  "Interest Earned",
  "Other",
];

const DEFAULT_CATEGORIES = [
  "Institute & Hostel Fees",
  "Mess / Food",
  "Laptop / Electronics",
  "Studio Supplies / Stationery",
  "Printing / Plotting",
  "Model-Making / Fabrication",
  "Books & References",
  "Local Field Visits",
  "Study Tour / Travel",
  "Home Travel",
  "Competitions / Workshops",
  "Health & Personal",
  "Local Commute",
  "Mobile / Internet & Subscriptions",
  "Internship / Training Expenses",
  "Portfolio / Placement",
  "Thesis Expenses",
  "Transport / Other",
];

// Extracted from College_Budget.xlsx — "Budget" sheet, values as shown (High scenario, 1.3x baked in per the Scenarios sheet).
// `variable: true` marks rows the Assumptions sheet prefixes "PERSONAL –" or "HOME TRAVEL –" — the ones the
// Scenarios sheet says the multiplier applies to. Everything else (official fees, mess, internship/portfolio/thesis
// one-offs) stays fixed regardless of scenario.
const BASE_SCENARIO_MULTIPLIER = 1.3; // the multiplier already baked into the sheet's shown numbers (its "High")
const SCENARIO_MULTIPLIERS = { Conservative: 0.75, Comfortable: 1, High: 1.25 };

const BUDGET_ROWS = [
  { head: "Institute & Hostel Fees", years: [120000, 100000, 100000, 100000, 100000], notes: "Year 1 includes one-time fees & refundable deposits", mapsTo: "Institute & Hostel Fees", group: "Fixed Cost", variable: false },
  { head: "Mess / Food", years: [50000, 50000, 60000, 60000, 80000], notes: "Paid to vendor per semester", mapsTo: "Mess / Food", group: "Fixed Cost", variable: false },
  { head: "Laptop + Mobile (one-time)", years: [200000, 0, 0, 0, 0], notes: "Bought in Year 1", mapsTo: "Laptop / Electronics", group: "Personal", variable: true },
  { head: "Laptop upgrade / repair", years: [0, 0, 5000, 5000, 5000], notes: "Optional buffer", mapsTo: "Laptop / Electronics", group: "Personal", variable: true },
  { head: "Studio Supplies / Stationery", years: [20000, 20000, 20000, 20000, 20000], notes: "Sheets, pens, tools, adhesives", mapsTo: "Studio Supplies / Stationery", group: "Study Materials / Printing", variable: true },
  { head: "Printing / Plotting", years: [20000, 30000, 30000, 30000, 40000], notes: "Submissions & juries", mapsTo: "Printing / Plotting", group: "Study Materials / Printing", variable: true },
  { head: "Model-Making / Fabrication", years: [10000, 10000, 10000, 20000, 20000], notes: "Foamboard / MDF / 3D print / laser", mapsTo: "Model-Making / Fabrication", group: "Study Materials / Printing", variable: true },
  { head: "Books & References", years: [10000, 10000, 10000, 10000, 10000], notes: "Purchased books / standards", mapsTo: "Books & References", group: "Study Materials / Printing", variable: true },
  { head: "Local Field Visits", years: [10000, 10000, 20000, 20000, 20000], notes: "Local trips with college", mapsTo: "Local Field Visits", group: "Study Tours / Internships", variable: true },
  { head: "Major Study Tour (Year 2)", years: [0, 20000, 0, 0, 0], notes: "When scheduled", mapsTo: "Study Tour / Travel", group: "Study Tours / Internships", variable: true },
  { head: "Major Study Tour (Year 4)", years: [0, 0, 0, 30000, 0], notes: "When scheduled", mapsTo: "Study Tour / Travel", group: "Study Tours / Internships", variable: true },
  { head: "Competitions / Workshops", years: [10000, 20000, 20000, 20000, 20000], notes: "Probable", mapsTo: "Competitions / Workshops", group: "Study Tours / Internships", variable: true },
  { head: "Health & Personal", years: [10000, 15000, 20000, 25000, 30000], notes: "Medicines / tests", mapsTo: "Health & Personal", group: "Personal", variable: true },
  { head: "Local Commute", years: [10000, 15000, 20000, 25000, 30000], notes: "Autos / cabs / buses", mapsTo: "Local Commute", group: "Personal", variable: true },
  { head: "Mobile / Internet & Subscriptions", years: [12000, 12000, 15000, 15000, 15000], notes: "Data plans / cloud", mapsTo: "Mobile / Internet & Subscriptions", group: "Personal", variable: true },
  { head: "Home Travel (round trips)", years: [24000, 40000, 40000, 40000, 40000], notes: "Adjust trips & per-trip cost", mapsTo: "Home Travel", group: "Personal", variable: true },
  { head: "Internship / Training (Year 4)", years: [0, 0, 0, 200000, 0], notes: "Travel + stay + local commute", mapsTo: "Internship / Training Expenses", group: "Study Tours / Internships", variable: false },
  { head: "Portfolio / Placement (Year 5)", years: [0, 0, 0, 0, 20000], notes: "Printing / portfolio", mapsTo: "Portfolio / Placement", group: "Study Materials / Printing", variable: false },
  { head: "Thesis Buffer (Year 5)", years: [0, 0, 0, 0, 20000], notes: "Thesis peak spend", mapsTo: "Thesis Expenses", group: "Study Materials / Printing", variable: false },
];

const GROUPS = ["Fixed Cost", "Study Materials / Printing", "Study Tours / Internships", "Personal"];

const CHART_COLORS = ["#3b82f6", "#ef4444", "#22c55e", "#f59e0b", "#a855f7", "#06b6d4", "#ec4899", "#84cc16", "#f97316", "#6366f1", "#14b8a6", "#eab308", "#f43f5e", "#8b5cf6", "#10b981", "#0ea5e9", "#d946ef", "#64748b"];

const CURRENCY_PRESETS = [
  { label: "INR (₹)", symbol: "₹", locale: "en-IN" },
  { label: "USD ($)", symbol: "$", locale: "en-US" },
  { label: "EUR (€)", symbol: "€", locale: "en-IE" },
  { label: "GBP (£)", symbol: "£", locale: "en-GB" },
];

const PAYMENT_METHODS = ["Cash", "Card", "UPI", "Other"];

const initialTransactions = [];

const STORAGE_KEY = "campcash-wallet-state";

function formatDate(iso) {
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" });
}

const ROLLOVER_CATEGORY = "Month Rollover";
const FD_CATEGORY = "Fixed Deposit";

function isTransferCategory(cat) {
  return cat === ROLLOVER_CATEGORY || cat === FD_CATEGORY;
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function monthKeyFromDate(iso) {
  return iso.slice(0, 7); // "YYYY-MM"
}

function lastDayOfMonthNum(year, monthNum1) {
  // monthNum1 is 1-indexed; day 0 of the following month is the last day of this one
  return new Date(year, monthNum1, 0).getDate();
}

function monthLabel(year, monthNum1) {
  return new Date(year, monthNum1 - 1, 1).toLocaleDateString("en-US", { month: "long", year: "numeric" });
}

function fdMaturityDate(startDateISO, tenureMonths) {
  const d = new Date(startDateISO + "T00:00:00");
  d.setMonth(d.getMonth() + tenureMonths);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function fdMaturityAmount(principal, ratePct, tenureMonths, compounding) {
  const years = tenureMonths / 12;
  const r = ratePct / 100;
  if (compounding === "simple") return principal * (1 + r * years);
  const n = { monthly: 12, quarterly: 4, annually: 1 }[compounding] || 4;
  return principal * Math.pow(1 + r / n, n * years);
}

function fdProgressPct(startDateISO, maturityDateISO) {
  const start = new Date(startDateISO + "T00:00:00").getTime();
  const end = new Date(maturityDateISO + "T00:00:00").getTime();
  const now = Date.now();
  if (now <= start) return 0;
  if (now >= end) return 100;
  return Math.round(((now - start) / (end - start)) * 100);
}

// Templates for a 3-month demo ledger, expressed relative to "today" so the sample data is always
// current whenever this is actually run — never tied to the date this code was written.
const SAMPLE_MONTH_TEMPLATES = [
  // two months ago
  [
    { day: 1, type: "income", category: "Parents / Family", amount: 5000, notes: "Monthly allowance" },
    { day: 3, type: "expense", category: "Mess / Food", amount: 1800, notes: "Mess bill", method: "Card" },
    { day: 5, type: "expense", category: "Local Commute", amount: 220, notes: "Auto rides", method: "Cash" },
    { day: 7, type: "expense", category: "Studio Supplies / Stationery", amount: 480, notes: "Sheets and pens", method: "Cash" },
    { day: 9, type: "expense", category: "Printing / Plotting", amount: 650, notes: "Submission prints", method: "UPI" },
    { day: 11, type: "expense", category: "Mobile / Internet & Subscriptions", amount: 399, notes: "Recharge", method: "UPI" },
    { day: 14, type: "income", category: "Refund", amount: 300, notes: "Refund from vendor" },
    { day: 16, type: "expense", category: "Health & Personal", amount: 220, notes: "Medicines", method: "Cash" },
    { day: 18, type: "expense", category: "Model-Making / Fabrication", amount: 750, notes: "Foam board + adhesive", method: "Card" },
    { day: 21, type: "expense", category: "Books & References", amount: 350, notes: "Reference book", method: "UPI" },
    { day: 24, type: "expense", category: "Local Field Visits", amount: 400, notes: "Site visit auto fare", method: "Cash" },
    { day: 27, type: "expense", category: "Transport / Other", amount: 180, notes: "Misc travel", method: "Cash" },
  ],
  // last month
  [
    { day: 1, type: "income", category: "Parents / Family", amount: 5000, notes: "Monthly allowance" },
    { day: 2, type: "income", category: "Scholarship", amount: 2000, notes: "Semester scholarship" },
    { day: 4, type: "expense", category: "Mess / Food", amount: 1900, notes: "Mess bill", method: "Card" },
    { day: 6, type: "expense", category: "Local Commute", amount: 260, notes: "Auto rides", method: "Cash" },
    { day: 8, type: "expense", category: "Competitions / Workshops", amount: 500, notes: "Workshop fee", method: "Card" },
    { day: 10, type: "expense", category: "Printing / Plotting", amount: 700, notes: "Model review prints", method: "UPI" },
    { day: 13, type: "expense", category: "Studio Supplies / Stationery", amount: 400, notes: "Craft materials", method: "Cash" },
    { day: 16, type: "expense", category: "Home Travel", amount: 800, notes: "Trip home", method: "Card" },
    { day: 19, type: "expense", category: "Mobile / Internet & Subscriptions", amount: 399, notes: "Recharge", method: "UPI" },
    { day: 22, type: "expense", category: "Health & Personal", amount: 300, notes: "Doctor visit", method: "Cash" },
    { day: 25, type: "expense", category: "Model-Making / Fabrication", amount: 650, notes: "3D print", method: "Card" },
    { day: 28, type: "expense", category: "Local Field Visits", amount: 420, notes: "Site visit", method: "Cash" },
  ],
  // this month (only entries up to today get included)
  [
    { day: 1, type: "income", category: "Parents / Family", amount: 5000, notes: "Monthly allowance" },
    { day: 3, type: "expense", category: "Mess / Food", amount: 1850, notes: "Mess bill", method: "Card" },
    { day: 5, type: "expense", category: "Local Commute", amount: 230, notes: "Auto rides", method: "Cash" },
    { day: 7, type: "expense", category: "Studio Supplies / Stationery", amount: 420, notes: "Sheets and pens", method: "Cash" },
    { day: 9, type: "expense", category: "Printing / Plotting", amount: 600, notes: "Submission prints", method: "UPI" },
    { day: 11, type: "expense", category: "Mobile / Internet & Subscriptions", amount: 399, notes: "Recharge", method: "UPI" },
    { day: 13, type: "expense", category: "Books & References", amount: 300, notes: "Reference book", method: "UPI" },
    { day: 20, type: "expense", category: "Competitions / Workshops", amount: 450, notes: "Workshop fee", method: "Card" },
    { day: 26, type: "expense", category: "Local Field Visits", amount: 380, notes: "Site visit auto fare", method: "Cash" },
  ],
];

function generateSampleDataEntries() {
  const now = new Date();
  const entries = [];
  for (let back = 2; back >= 0; back--) {
    const ref = new Date(now.getFullYear(), now.getMonth() - back, 1);
    const y = ref.getFullYear();
    const m = ref.getMonth() + 1; // 1-indexed
    const template = SAMPLE_MONTH_TEMPLATES[2 - back];
    template.forEach(entry => {
      if (back === 0 && entry.day > now.getDate()) return; // don't backdate into the future
      const lastDay = lastDayOfMonthNum(y, m);
      const day = Math.min(entry.day, lastDay);
      const date = `${y}-${pad2(m)}-${pad2(day)}`;
      entries.push({
        date, type: entry.type, amount: entry.amount, category: entry.category, notes: entry.notes,
        ...(entry.type === "expense" ? { paymentMethod: entry.method || "Other" } : {}),
      });
    });
  }
  return entries;
}

function parseCSV(text) {
  const rows = [];
  let cur = [], field = "", inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ",") { cur.push(field); field = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        cur.push(field); field = "";
        rows.push(cur); cur = [];
      } else field += c;
    }
  }
  if (field.length || cur.length) { cur.push(field); rows.push(cur); }
  return rows.filter(r => r.length > 1 || (r.length === 1 && r[0].trim() !== ""));
}

export default function CampCash() {
  const [dark, setDark] = useState(false);
  const [activeView, setActiveView] = useState("home");
  const [transactions, setTransactions] = useState(initialTransactions);
  const [nextId, setNextId] = useState(initialTransactions.length + 1);
  const [search, setSearch] = useState("");
  const [filterCategory, setFilterCategory] = useState("All");
  const [filterOpen, setFilterOpen] = useState(false);
  const [budgetYear, setBudgetYear] = useState(1);
  const [scenario, setScenario] = useState("High");

  const [sources, setSources] = useState(DEFAULT_SOURCES);
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES);
  const [newSourceInput, setNewSourceInput] = useState("");
  const [newCategoryInput, setNewCategoryInput] = useState("");

  const [currency, setCurrency] = useState(CURRENCY_PRESETS[0]);
  const [openingBalance, setOpeningBalance] = useState(0);
  const [openingBalanceInput, setOpeningBalanceInput] = useState("0");
  const [accountName, setAccountName] = useState("");
  const [accountEmail, setAccountEmail] = useState("");
  const [resetConfirming, setResetConfirming] = useState(false);
  const [settingsMessage, setSettingsMessage] = useState("");
  const [notifOpen, setNotifOpen] = useState(false);
  const [seenNotifIds, setSeenNotifIds] = useState(() => new Set());
  const [menuOpen, setMenuOpen] = useState(false);
  const [dataLoaded, setDataLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState("idle"); // idle | saving | saved | error
  const [lastRolloverMonth, setLastRolloverMonth] = useState(null); // "YYYY-MM" — months up to (not including) this have been closed out

  const [periodMode, setPeriodMode] = useState("month"); // "month" | "lifetime"
  const [periodMonthKey, setPeriodMonthKey] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}`;
  });
  const [periodOpen, setPeriodOpen] = useState(false);

  const [fds, setFds] = useState([]);
  const [nextFdId, setNextFdId] = useState(1);
  const [fdLabel, setFdLabel] = useState("");
  const [fdPrincipal, setFdPrincipal] = useState("");
  const [fdRate, setFdRate] = useState("");
  const [fdStartDate, setFdStartDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [fdTenure, setFdTenure] = useState("");
  const [fdCompounding, setFdCompounding] = useState("quarterly");
  const [fdDebitOnCreate, setFdDebitOnCreate] = useState(true);
  const [fdError, setFdError] = useState("");
  const [closingFdId, setClosingFdId] = useState(null);
  const [closingDraft, setClosingDraft] = useState(null); // { status, credit, amount }

  function formatMoney(amount) {
    return currency.symbol + Math.round(amount).toLocaleString(currency.locale);
  }

  const [addAmount, setAddAmount] = useState("");
  const [addSource, setAddSource] = useState("");
  const [addNotes, setAddNotes] = useState("");

  const [spendAmount, setSpendAmount] = useState("");
  const [spendCategory, setSpendCategory] = useState("");
  const [spendNotes, setSpendNotes] = useState("");
  const [spendMethod, setSpendMethod] = useState("");

  const [error, setError] = useState("");

  const [sortKey, setSortKey] = useState(null); // null | "date" | "amount"
  const [sortDir, setSortDir] = useState("desc"); // "asc" | "desc"

  const [lastDeleted, setLastDeleted] = useState(null);
  const undoTimeoutRef = useRef(null);

  const jsonFileInputRef = useRef(null);
  const csvFileInputRef = useRef(null);
  const [restoreConfirming, setRestoreConfirming] = useState(false);
  const [pendingRestore, setPendingRestore] = useState(null);
  const [restoreSource, setRestoreSource] = useState("file"); // "file" | "cloud" — customizes the confirm banner wording

  const [githubToken, setGithubToken] = useState("");
  const [gistId, setGistId] = useState("");
  const [cloudConnected, setCloudConnected] = useState(false);
  const [cloudSyncStatus, setCloudSyncStatus] = useState("disabled"); // disabled | connecting | syncing | synced | error
  const [cloudLastSyncedAt, setCloudLastSyncedAt] = useState(null);

  // Load any previously saved wallet on first mount
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw && !cancelled) {
          const saved = JSON.parse(raw);
          if (Array.isArray(saved.transactions)) setTransactions(saved.transactions);
          if (typeof saved.nextId === "number") setNextId(saved.nextId);
          if (Array.isArray(saved.sources)) setSources(saved.sources);
          if (Array.isArray(saved.categories)) setCategories(saved.categories);
          if (saved.currency) setCurrency(saved.currency);
          if (typeof saved.openingBalance === "number") {
            setOpeningBalance(saved.openingBalance);
            setOpeningBalanceInput(String(saved.openingBalance));
          }
          if (typeof saved.accountName === "string") setAccountName(saved.accountName);
          if (typeof saved.accountEmail === "string") setAccountEmail(saved.accountEmail);
          if (typeof saved.dark === "boolean") setDark(saved.dark);
          if (typeof saved.scenario === "string") setScenario(saved.scenario);
          if (typeof saved.lastRolloverMonth === "string") setLastRolloverMonth(saved.lastRolloverMonth);
          if (Array.isArray(saved.fds)) setFds(saved.fds);
          if (typeof saved.nextFdId === "number") setNextFdId(saved.nextFdId);
          if (typeof saved.githubToken === "string") setGithubToken(saved.githubToken);
          if (typeof saved.gistId === "string") setGistId(saved.gistId);
          if (typeof saved.cloudConnected === "boolean") setCloudConnected(saved.cloudConnected);
          if (typeof saved.cloudLastSyncedAt === "string") setCloudLastSyncedAt(saved.cloudLastSyncedAt);
        }
      } catch (e) {
        // Nothing saved yet — that's a normal first run, not an error.
      } finally {
        if (!cancelled) setDataLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  // Autosave, debounced — only committed data (not in-progress form drafts)
  useEffect(() => {
    if (!dataLoaded) return;
    setSaveStatus("saving");
    const handle = setTimeout(() => {
      try {
        const payload = {
          transactions, nextId, sources, categories, currency,
          openingBalance, accountName, accountEmail, dark, scenario, lastRolloverMonth,
          fds, nextFdId, githubToken, gistId, cloudConnected, cloudLastSyncedAt,
          savedAt: new Date().toISOString(),
        };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
        setSaveStatus("saved");
      } catch (e) {
        setSaveStatus("error");
      }
    }, 500);
    return () => clearTimeout(handle);
  }, [transactions, nextId, sources, categories, currency, openingBalance, accountName, accountEmail, dark, scenario, lastRolloverMonth, fds, nextFdId, githubToken, gistId, cloudConnected, cloudLastSyncedAt, dataLoaded]);


  // Month-end rollover: for every calendar month that has fully passed since we last checked,
  // record an explicit "closing balance" entry on its last day and a matching "opening balance"
  // entry on the 1st of the next month — same amount, so the real balance never changes, but
  // the ledger shows a visible carry-forward line the way a bank statement would.
  useEffect(() => {
    if (!dataLoaded) return;
    const now = new Date();
    const currentKey = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}`;

    if (transactions.length === 0) {
      if (!lastRolloverMonth) setLastRolloverMonth(currentKey);
      return;
    }

    const earliestDate = transactions.reduce((min, t) => (t.date < min ? t.date : min), transactions[0].date);
    const earliestKey = monthKeyFromDate(earliestDate);
    let startKey = (lastRolloverMonth && lastRolloverMonth < earliestKey) ? lastRolloverMonth : earliestKey;

    if (startKey >= currentKey) {
      if (!lastRolloverMonth) setLastRolloverMonth(currentKey);
      return;
    }

    const newEntries = [];
    let idCounter = nextId;
    let [y, m] = startKey.split("-").map(Number);
    const [cy, cm] = currentKey.split("-").map(Number);
    let safety = 0;

    while ((y < cy || (y === cy && m < cm)) && safety < 36) {
      safety++;
      const lastDay = lastDayOfMonthNum(y, m);
      const outDate = `${y}-${pad2(m)}-${pad2(lastDay)}`;
      const nextM = m === 12 ? 1 : m + 1;
      const nextY = m === 12 ? y + 1 : y;
      const inDate = `${nextY}-${pad2(nextM)}-01`;

      const closingBalance = openingBalance + transactions
        .filter(t => t.date <= outDate)
        .reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);

      if (closingBalance > 0.5) {
        const rounded = Math.round(closingBalance);
        newEntries.push({
          id: idCounter++, date: outDate, type: "expense", amount: rounded,
          category: ROLLOVER_CATEGORY, notes: `Balance carried forward to ${monthLabel(nextY, nextM)}`,
        });
        newEntries.push({
          id: idCounter++, date: inDate, type: "income", amount: rounded,
          category: ROLLOVER_CATEGORY, notes: `Balance brought forward from ${monthLabel(y, m)}`,
        });
      }

      y = nextY; m = nextM;
    }

    if (newEntries.length) {
      setTransactions(prev => [...prev, ...newEntries]);
      setNextId(idCounter);
    }
    setLastRolloverMonth(currentKey);
  }, [dataLoaded, transactions, nextId, openingBalance, lastRolloverMonth]);

  // Cloud sync (GitHub Gist) — debounced auto-push whenever connected and something changes
  useEffect(() => {
    if (!dataLoaded || !cloudConnected || !gistId.trim() || !githubToken.trim()) return;
    setCloudSyncStatus("syncing");
    const handle = setTimeout(async () => {
      try {
        const payload = buildFullBackupPayload();
        await githubPushGist(payload);
        setCloudLastSyncedAt(payload.savedAt);
        setCloudSyncStatus("synced");
      } catch (e) {
        setCloudSyncStatus("error");
      }
    }, 1200);
    return () => clearTimeout(handle);
  }, [transactions, nextId, sources, categories, currency, openingBalance, accountName, accountEmail, scenario, lastRolloverMonth, fds, nextFdId, cloudConnected, gistId, githubToken, dataLoaded]);

  // On load, if a previous session already connected cloud sync, check whether another device
  // pushed something newer and offer to load it — never silently overwrites local work.
  useEffect(() => {
    if (!dataLoaded || !cloudConnected || !gistId.trim() || !githubToken.trim()) return;
    let cancelled = false;
    (async () => {
      try {
        const remote = await githubPullGist();
        if (cancelled || !remote || !Array.isArray(remote.transactions)) return;
        if (remote.savedAt && remote.savedAt !== cloudLastSyncedAt) {
          setRestoreSource("cloud");
          setPendingRestore(remote);
          setRestoreConfirming(true);
        }
      } catch (e) {
        setCloudSyncStatus("error");
      }
    })();
    return () => { cancelled = true; };
  }, [dataLoaded]);

  const chronological = useMemo(
    () => [...transactions].sort((a, b) => a.date === b.date ? a.id - b.id : a.date.localeCompare(b.date)),
    [transactions]
  );

  const withBalance = useMemo(() => {
    let bal = openingBalance;
    return chronological.map((t) => {
      bal += t.type === "income" ? t.amount : -t.amount;
      return { ...t, balanceAfter: bal };
    });
  }, [chronological, openingBalance]);

  const balance = withBalance.length ? withBalance[withBalance.length - 1].balanceAfter : openingBalance;

  const isCurrentMonth = useMemo(() => {
    const now = new Date();
    return periodMonthKey === `${now.getFullYear()}-${pad2(now.getMonth() + 1)}`;
  }, [periodMonthKey]);

  const periodLabel = periodMode === "lifetime"
    ? "Lifetime"
    : (isCurrentMonth ? "This Month" : (() => { const [y, m] = periodMonthKey.split("-").map(Number); return monthLabel(y, m); })());

  const periodIn = transactions
    .filter(t => t.type === "income" && !isTransferCategory(t.category) && (periodMode === "lifetime" || monthKeyFromDate(t.date) === periodMonthKey))
    .reduce((s, t) => s + t.amount, 0);
  const periodOut = transactions
    .filter(t => t.type === "expense" && !isTransferCategory(t.category) && (periodMode === "lifetime" || monthKeyFromDate(t.date) === periodMonthKey))
    .reduce((s, t) => s + t.amount, 0);

  const displayList = useMemo(() => {
    let list = [...withBalance].reverse();
    if (periodMode === "month") {
      list = list.filter(t => monthKeyFromDate(t.date) === periodMonthKey);
    }
    if (filterCategory !== "All") {
      list = list.filter(t => t.category === filterCategory);
    }
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      list = list.filter(t =>
        (t.notes || "").toLowerCase().includes(q) ||
        (t.category || "").toLowerCase().includes(q)
      );
    }
    if (sortKey === "date") {
      list = [...list].sort((a, b) => sortDir === "asc" ? a.date.localeCompare(b.date) : b.date.localeCompare(a.date));
    } else if (sortKey === "amount") {
      list = [...list].sort((a, b) => sortDir === "asc" ? a.amount - b.amount : b.amount - a.amount);
    }
    return list;
  }, [withBalance, search, filterCategory, sortKey, sortDir, periodMode, periodMonthKey]);

  function toggleSort(key) {
    if (sortKey === key) {
      setSortDir(d => d === "asc" ? "desc" : "asc");
    } else {
      setSortKey(key);
      setSortDir("desc");
    }
  }

  // Actual spend per category, for comparing against the budget plan
  const actualByCategory = useMemo(() => {
    const map = {};
    transactions.filter(t => t.type === "expense" && !isTransferCategory(t.category)).forEach(t => {
      map[t.category] = (map[t.category] || 0) + t.amount;
    });
    return map;
  }, [transactions]);

  // Actual income per source
  const actualBySource = useMemo(() => {
    const map = {};
    transactions.filter(t => t.type === "income" && !isTransferCategory(t.category)).forEach(t => {
      map[t.category] = (map[t.category] || 0) + t.amount;
    });
    return map;
  }, [transactions]);

  const spendingChartData = useMemo(
    () => Object.entries(actualByCategory)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value),
    [actualByCategory]
  );

  const incomeChartData = useMemo(
    () => Object.entries(actualBySource)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value),
    [actualBySource]
  );

  // Scoped to whatever's selected in the sidebar's time-period control — used only by the
  // Analytics tab's category/source breakdowns. The Balance Trend chart stays full history,
  // and the Budget Plan / notifications keep using the lifetime totals above.
  const periodTransactions = useMemo(
    () => transactions.filter(t => periodMode === "lifetime" || monthKeyFromDate(t.date) === periodMonthKey),
    [transactions, periodMode, periodMonthKey]
  );

  const periodSpendingChartData = useMemo(() => {
    const map = {};
    periodTransactions.filter(t => t.type === "expense" && !isTransferCategory(t.category)).forEach(t => {
      map[t.category] = (map[t.category] || 0) + t.amount;
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [periodTransactions]);

  const periodIncomeChartData = useMemo(() => {
    const map = {};
    periodTransactions.filter(t => t.type === "income" && !isTransferCategory(t.category)).forEach(t => {
      map[t.category] = (map[t.category] || 0) + t.amount;
    });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [periodTransactions]);

  const balanceTrendData = useMemo(
    () => [
      { label: "Start", balance: openingBalance },
      ...withBalance.map((t, i) => ({
        label: `${formatDate(t.date)}${i > 0 && withBalance[i - 1].date === t.date ? " *" : ""}`,
        balance: t.balanceAfter,
      })),
    ],
    [withBalance, openingBalance]
  );

  // Rescale every row to the selected scenario: variable rows go back to their un-multiplied
  // base (÷1.3, the multiplier baked into the sheet's shown numbers) then apply the new multiplier.
  // Fixed rows (official fees, mess, internship/portfolio/thesis one-offs) never change.
  const scaledRows = useMemo(() => {
    const mult = SCENARIO_MULTIPLIERS[scenario];
    return BUDGET_ROWS.map(row => ({
      ...row,
      scaledYears: row.years.map(v => row.variable ? Math.round((v / BASE_SCENARIO_MULTIPLIER) * mult) : v),
    }));
  }, [scenario]);

  const yearTotals = useMemo(() => {
    const totals = [0, 0, 0, 0, 0];
    scaledRows.forEach(row => row.scaledYears.forEach((v, i) => { totals[i] += v; }));
    return totals;
  }, [scaledRows]);

  const fiveYearTotal = yearTotals.reduce((a, b) => a + b, 0);
  const monthlySetAside = fiveYearTotal / 60;

  const groupTotals = useMemo(() => {
    return GROUPS.map(group => ({
      group,
      years: [0, 1, 2, 3, 4].map(i =>
        scaledRows.filter(r => r.group === group).reduce((s, r) => s + r.scaledYears[i], 0)
      ),
    }));
  }, [scaledRows]);

  const budgetForYear = useMemo(() => {
    const idx = budgetYear - 1;
    const byCategory = {};
    scaledRows.forEach(row => {
      byCategory[row.mapsTo] = (byCategory[row.mapsTo] || 0) + row.scaledYears[idx];
    });
    return byCategory;
  }, [scaledRows, budgetYear]);

  // Always compares against Year 1 of the current scenario, regardless of which year is
  // being browsed on the Budget Plan tab — this is about today's actual spending, not planning.
  const budgetForYearOne = useMemo(() => {
    const byCategory = {};
    scaledRows.forEach(row => {
      byCategory[row.mapsTo] = (byCategory[row.mapsTo] || 0) + row.scaledYears[0];
    });
    return byCategory;
  }, [scaledRows]);

  const notifications = useMemo(() => {
    const list = [];
    categories.forEach(cat => {
      const budgeted = budgetForYearOne[cat] || 0;
      const actual = actualByCategory[cat] || 0;
      if (budgeted > 0 && actual > budgeted) {
        list.push({
          id: `over-${cat}`,
          tone: "warning",
          text: `${cat} is over budget — ${formatMoney(actual)} logged against ${formatMoney(budgeted)} planned for Year 1 (${scenario}).`,
        });
      }
    });
    if (transactions.length > 0 && balance < monthlySetAside) {
      list.push({
        id: "low-balance",
        tone: "warning",
        text: `Your balance (${formatMoney(balance)}) is below your suggested monthly set-aside of ${formatMoney(monthlySetAside)}.`,
      });
    }
    if (transactions.length === 0) {
      list.push({ id: "welcome", tone: "info", text: "Log your first transaction on Home to start tracking against your budget." });
    }
    const todayISO = new Date().toISOString().slice(0, 10);
    fds.filter(fd => fd.status === "active").forEach(fd => {
      const maturityDate = fdMaturityDate(fd.startDate, fd.tenureMonths);
      if (maturityDate <= todayISO) {
        list.push({
          id: `fd-matured-${fd.id}`,
          tone: "info",
          text: `Your FD "${fd.label}" matured on ${formatDate(maturityDate)} — mark it as matured on the Fixed Deposits tab to update your balance.`,
        });
      }
    });
    return list;
  }, [categories, budgetForYearOne, actualByCategory, balance, monthlySetAside, transactions, scenario, fds]);

  const unseenNotifCount = notifications.filter(n => !seenNotifIds.has(n.id)).length;

  function toggleNotifPanel() {
    setNotifOpen(o => {
      const next = !o;
      if (next) {
        setSeenNotifIds(prev => {
          const s = new Set(prev);
          notifications.forEach(n => s.add(n.id));
          return s;
        });
      }
      return next;
    });
    setMenuOpen(false);
  }

  function addTransaction({ type, amount, category, notes, paymentMethod, dateOverride, idOverride }) {
    const amt = parseFloat(amount);
    if (!amt || amt <= 0) {
      setError(type === "income" ? "Enter a valid amount to add." : "Enter a valid amount to spend.");
      return false;
    }
    if (type === "expense" && amt > balance) {
      setError("That's more than your current balance.");
      return false;
    }
    setError("");
    const today = new Date().toISOString().slice(0, 10);
    const usedId = idOverride || nextId;
    setTransactions(prev => [
      ...prev,
      {
        id: usedId,
        date: dateOverride || today,
        type, amount: amt,
        category: category || "Other",
        notes: notes || "",
        ...(type === "expense" && !isTransferCategory(category) ? { paymentMethod: paymentMethod || "Other" } : {}),
      },
    ]);
    if (!idOverride) setNextId(id => id + 1);
    return usedId;
  }

  function handleDeposit() {
    if (addTransaction({ type: "income", amount: addAmount, category: addSource, notes: addNotes })) {
      setAddAmount(""); setAddSource(""); setAddNotes("");
    }
  }

  function handleSpend() {
    if (addTransaction({ type: "expense", amount: spendAmount, category: spendCategory, notes: spendNotes, paymentMethod: spendMethod })) {
      setSpendAmount(""); setSpendCategory(""); setSpendNotes(""); setSpendMethod("");
    }
  }

  function handleQuickAdd(type, amt) {
    addTransaction({ type, amount: amt, category: "Other", notes: "Quick add", paymentMethod: "Other" });
  }

  function deleteTransaction(id) {
    const tx = transactions.find(t => t.id === id);
    setTransactions(prev => prev.filter(t => t.id !== id));
    if (tx) {
      if (undoTimeoutRef.current) clearTimeout(undoTimeoutRef.current);
      setLastDeleted(tx);
      undoTimeoutRef.current = setTimeout(() => setLastDeleted(null), 6000);
    }
  }

  function undoDelete() {
    if (!lastDeleted) return;
    setTransactions(prev => [...prev, lastDeleted]);
    setLastDeleted(null);
    if (undoTimeoutRef.current) clearTimeout(undoTimeoutRef.current);
  }

  function loadSampleData() {
    const entries = generateSampleDataEntries();
    let idCounter = nextId;
    const withIds = entries.map(e => ({ id: idCounter++, ...e }));
    setTransactions(prev => [...prev, ...withIds]);
    setNextId(idCounter);
    setSettingsMessage(`Loaded ${withIds.length} sample transactions across the last 3 months.`);
  }

  function addFD() {
    const principal = parseFloat(fdPrincipal);
    const rate = parseFloat(fdRate);
    const tenure = parseInt(fdTenure, 10);
    if (!fdLabel.trim() || !principal || principal <= 0 || !rate || rate <= 0 || !tenure || tenure <= 0 || !fdStartDate) {
      setFdError("Fill in all fields with valid values.");
      return;
    }
    let debitedTxId = null;
    if (fdDebitOnCreate) {
      const usedId = addTransaction({
        type: "expense", amount: principal, category: FD_CATEGORY,
        notes: `Opened FD — ${fdLabel.trim()}`, dateOverride: fdStartDate,
      });
      if (!usedId) {
        setFdError("That's more than your current balance — uncheck \"debit from wallet\" or adjust the amount.");
        return;
      }
      debitedTxId = usedId;
    }
    setFdError("");
    setFds(prev => [...prev, {
      id: nextFdId, label: fdLabel.trim(), principal, rate, startDate: fdStartDate,
      tenureMonths: tenure, compounding: fdCompounding, status: "active",
      debited: fdDebitOnCreate, debitedTxId,
    }]);
    setNextFdId(id => id + 1);
    setFdLabel(""); setFdPrincipal(""); setFdRate(""); setFdTenure("");
    setFdStartDate(new Date().toISOString().slice(0, 10)); setFdCompounding("quarterly"); setFdDebitOnCreate(true);
  }

  function startClosingFd(id, status, defaultAmount) {
    setFdError("");
    setClosingFdId(id);
    setClosingDraft({ status, credit: true, amount: String(Math.round(defaultAmount)) });
  }

  function cancelClosingFd() {
    setClosingFdId(null);
    setClosingDraft(null);
  }

  function confirmCloseFd() {
    const fd = fds.find(f => f.id === closingFdId);
    if (!fd || !closingDraft) return;
    const payout = parseFloat(closingDraft.amount);
    if (!payout || payout <= 0) {
      setFdError("Enter a valid payout amount.");
      return;
    }
    let creditedTxId = null;
    if (closingDraft.credit) {
      creditedTxId = addTransaction({
        type: "income", amount: payout, category: FD_CATEGORY,
        notes: `${closingDraft.status === "matured" ? "FD matured" : "FD withdrawn early"} — ${fd.label}`,
      });
    }
    setFds(prev => prev.map(f => f.id === closingFdId
      ? { ...f, status: closingDraft.status, credited: closingDraft.credit, creditedTxId, closedDate: new Date().toISOString().slice(0, 10), payoutAmount: payout }
      : f));
    setFdError("");
    setClosingFdId(null);
    setClosingDraft(null);
  }

  function deleteFd(id) {
    setFds(prev => prev.filter(f => f.id !== id));
    if (closingFdId === id) { setClosingFdId(null); setClosingDraft(null); }
  }

  const fdsWithComputed = useMemo(() => fds.map(fd => {
    const maturityDate = fdMaturityDate(fd.startDate, fd.tenureMonths);
    const maturityAmount = fdMaturityAmount(fd.principal, fd.rate, fd.tenureMonths, fd.compounding);
    return {
      ...fd,
      maturityDate,
      maturityAmount,
      expectedInterest: maturityAmount - fd.principal,
      progressPct: fdProgressPct(fd.startDate, maturityDate),
      isOverdue: fd.status === "active" && maturityDate <= new Date().toISOString().slice(0, 10),
    };
  }).sort((a, b) => a.maturityDate.localeCompare(b.maturityDate)), [fds]);

  const activeFds = fdsWithComputed.filter(fd => fd.status === "active");
  const closedFds = fdsWithComputed.filter(fd => fd.status !== "active");
  const fdTotalPrincipal = activeFds.reduce((s, fd) => s + fd.principal, 0);
  const fdTotalMaturity = activeFds.reduce((s, fd) => s + fd.maturityAmount, 0);
  const fdTotalInterest = fdTotalMaturity - fdTotalPrincipal;


  function handleResetAll() {
    setTransactions([]);
    setNextId(1);
    setFilterCategory("All");
    setSearch("");
    setResetConfirming(false);
    setSettingsMessage("All transactions cleared.");
  }

  function addSourceHeader() {
    const name = newSourceInput.trim();
    if (!name || sources.includes(name)) return;
    setSources(prev => [...prev, name]);
    setNewSourceInput("");
  }

  function removeSourceHeader(name) {
    setSources(prev => prev.filter(s => s !== name));
  }

  function addCategoryHeader() {
    const name = newCategoryInput.trim();
    if (!name || categories.includes(name)) return;
    setCategories(prev => [...prev, name]);
    setNewCategoryInput("");
  }

  function removeCategoryHeader(name) {
    setCategories(prev => prev.filter(c => c !== name));
  }

  function applyOpeningBalance() {
    const v = parseFloat(openingBalanceInput);
    setOpeningBalance(isNaN(v) ? 0 : v);
    setSettingsMessage("Opening balance updated.");
  }

  function buildFullBackupPayload() {
    return {
      savedAt: new Date().toISOString(),
      account: { name: accountName, email: accountEmail },
      currency,
      openingBalance,
      sources,
      categories,
      transactions,
      nextId,
      scenario,
      lastRolloverMonth,
      fds,
      nextFdId,
    };
  }

  function exportJSONBackup() {
    const payload = buildFullBackupPayload();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `campcash-backup-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setSettingsMessage("JSON backup downloaded.");
  }

  const GIST_FILENAME = "campcash-wallet.json";

  async function githubRequest(url, options = {}) {
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${githubToken.trim()}`,
        Accept: "application/vnd.github+json",
        ...(options.headers || {}),
      },
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`GitHub API ${res.status}: ${body.slice(0, 200)}`);
    }
    return res.json();
  }

  async function githubCreateGist(payload) {
    const data = await githubRequest("https://api.github.com/gists", {
      method: "POST",
      body: JSON.stringify({
        description: "CampCash wallet sync — do not delete",
        public: false,
        files: { [GIST_FILENAME]: { content: JSON.stringify(payload, null, 2) } },
      }),
    });
    return data.id;
  }

  async function githubPushGist(payload) {
    await githubRequest(`https://api.github.com/gists/${gistId.trim()}`, {
      method: "PATCH",
      body: JSON.stringify({
        files: { [GIST_FILENAME]: { content: JSON.stringify(payload, null, 2) } },
      }),
    });
  }

  async function githubPullGist() {
    const data = await githubRequest(`https://api.github.com/gists/${gistId.trim()}`);
    let file = data.files && data.files[GIST_FILENAME];
    if (!file) return null;
    let content = file.content;
    if (file.truncated && file.raw_url) {
      const res = await fetch(file.raw_url);
      content = await res.text();
    }
    try {
      return JSON.parse(content);
    } catch (e) {
      return null;
    }
  }

  async function connectCloudSync() {
    if (!githubToken.trim()) {
      setSettingsMessage("Enter your GitHub personal access token first.");
      return;
    }
    setCloudSyncStatus("connecting");
    try {
      if (!gistId.trim()) {
        // First device — create a fresh gist to hold the wallet
        const payload = buildFullBackupPayload();
        const newId = await githubCreateGist(payload);
        setGistId(newId);
        setCloudLastSyncedAt(payload.savedAt);
        setCloudConnected(true);
        setCloudSyncStatus("synced");
        setSettingsMessage(`Created a new sync gist (ID: ${newId}) — copy this ID into any other device to link it.`);
      } else {
        // Joining an existing gist from another device
        const remote = await githubPullGist();
        if (remote && Array.isArray(remote.transactions)) {
          setRestoreSource("cloud");
          setPendingRestore(remote);
          setRestoreConfirming(true);
          setCloudSyncStatus("synced");
        } else {
          const payload = buildFullBackupPayload();
          await githubPushGist(payload);
          setCloudLastSyncedAt(payload.savedAt);
          setCloudSyncStatus("synced");
          setSettingsMessage("Connected — this gist was empty, so your current wallet was pushed to it.");
        }
        setCloudConnected(true);
      }
    } catch (e) {
      setCloudSyncStatus("error");
      setSettingsMessage("Couldn't reach GitHub — check the token and Gist ID and try again.");
    }
  }

  function disconnectCloudSync() {
    setCloudConnected(false);
    setCloudSyncStatus("disabled");
    setSettingsMessage("Cloud sync disconnected. Your data stays local — the gist itself is untouched.");
  }

  async function manualSyncNow() {
    if (!cloudConnected) return;
    setCloudSyncStatus("syncing");
    try {
      const payload = buildFullBackupPayload();
      await githubPushGist(payload);
      setCloudLastSyncedAt(payload.savedAt);
      setCloudSyncStatus("synced");
      setSettingsMessage("Synced to cloud.");
    } catch (e) {
      setCloudSyncStatus("error");
    }
  }

  function csvEscape(val) {
    const s = String(val ?? "");
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function exportCSV() {
    const header = ["Date", "Type", "Category", "Payment Method", "Amount", "Balance After", "Notes"];
    const rows = withBalance.map(t => [
      t.date,
      t.type === "income" ? "Income" : "Expense",
      t.category,
      t.paymentMethod || "",
      t.amount,
      t.balanceAfter,
      t.notes || "",
    ]);
    const csvText = [header, ...rows].map(r => r.map(csvEscape).join(",")).join("\r\n");
    const blob = new Blob([csvText], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `campcash-ledger-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setSettingsMessage("CSV exported.");
  }

  function exportExcel() {
    const wb = XLSX.utils.book_new();

    // --- Ledger ---
    const ledgerRows = withBalance.map(t => ({
      Date: t.date,
      Type: t.type === "income" ? "Income" : "Expense",
      Category: t.category,
      "Payment Method": t.paymentMethod || "",
      Amount: t.type === "income" ? t.amount : -t.amount,
      "Balance After": t.balanceAfter,
      Notes: t.notes || "",
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ledgerRows), "Ledger");

    // --- Pivot: Spending by Category ---
    const totalSpend = Object.values(actualByCategory).reduce((a, b) => a + b, 0);
    const spendPivotRows = spendingChartData.map(c => ({
      Category: c.name,
      "Total Spent": c.value,
      Transactions: transactions.filter(t => t.type === "expense" && t.category === c.name).length,
      "% of Total Spend": totalSpend ? Math.round((c.value / totalSpend) * 1000) / 10 : 0,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(spendPivotRows), "Pivot - Spend by Category");

    // --- Pivot: Income by Source ---
    const totalIncome = Object.values(actualBySource).reduce((a, b) => a + b, 0);
    const incomePivotRows = incomeChartData.map(c => ({
      Source: c.name,
      "Total Received": c.value,
      Transactions: transactions.filter(t => t.type === "income" && t.category === c.name).length,
      "% of Total Income": totalIncome ? Math.round((c.value / totalIncome) * 1000) / 10 : 0,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(incomePivotRows), "Pivot - Income by Source");

    // --- Pivot: Payment Method ---
    const methodMap = {};
    transactions.filter(t => t.type === "expense" && !isTransferCategory(t.category)).forEach(t => {
      const key = t.paymentMethod || "Other";
      if (!methodMap[key]) methodMap[key] = { total: 0, count: 0 };
      methodMap[key].total += t.amount;
      methodMap[key].count += 1;
    });
    const methodPivotRows = Object.entries(methodMap)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([method, v]) => ({
        "Payment Method": method,
        "Total Spent": v.total,
        Transactions: v.count,
        "% of Total Spend": totalSpend ? Math.round((v.total / totalSpend) * 1000) / 10 : 0,
      }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(methodPivotRows), "Pivot - Payment Method");

    // --- Pivot: Monthly Summary ---
    const monthMap = {};
    withBalance.forEach(t => {
      const key = monthKeyFromDate(t.date);
      if (!monthMap[key]) monthMap[key] = { income: 0, expense: 0, closingBalance: t.balanceAfter };
      if (!isTransferCategory(t.category)) {
        if (t.type === "income") monthMap[key].income += t.amount;
        else monthMap[key].expense += t.amount;
      }
      monthMap[key].closingBalance = t.balanceAfter; // last one wins, in chronological order
    });
    const monthlyPivotRows = Object.keys(monthMap).sort().map(key => {
      const [y, m] = key.split("-").map(Number);
      return {
        Month: monthLabel(y, m),
        "Money In": monthMap[key].income,
        "Money Out": monthMap[key].expense,
        Net: monthMap[key].income - monthMap[key].expense,
        "Closing Balance": monthMap[key].closingBalance,
      };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthlyPivotRows), "Pivot - Monthly Summary");

    // --- Budget vs Actual (Year 1, current scenario) ---
    const budgetRows = categories.filter(c => c !== "Transport / Other").map(cat => {
      const budgeted = budgetForYearOne[cat] || 0;
      const actual = actualByCategory[cat] || 0;
      return {
        Category: cat,
        [`Budgeted (Year 1, ${scenario})`]: budgeted,
        "Actual Logged": actual,
        Variance: budgeted - actual,
        "% Used": budgeted ? Math.round((actual / budgeted) * 1000) / 10 : null,
      };
    }).filter(r => r["Actual Logged"] > 0 || r[`Budgeted (Year 1, ${scenario})`] > 0);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(budgetRows), "Budget vs Actual");

    // --- Summary ---
    const summaryRows = [
      { Metric: "Account", Value: accountName || "(not set)" },
      { Metric: "Currency", Value: currency.label },
      { Metric: "Exported On", Value: new Date().toLocaleString(currency.locale) },
      { Metric: "Current Balance", Value: balance },
      { Metric: `${periodLabel} — Money In`, Value: periodIn },
      { Metric: `${periodLabel} — Money Out`, Value: periodOut },
      { Metric: `${periodLabel} — Net`, Value: periodIn - periodOut },
      { Metric: "Scenario", Value: scenario },
      { Metric: "5-Year Budget Total", Value: fiveYearTotal },
      { Metric: "Suggested Monthly Set-Aside", Value: monthlySetAside },
      { Metric: "Total Transactions Logged", Value: transactions.length },
      { Metric: "Active Fixed Deposits", Value: activeFds.length },
      { Metric: "FD Principal Locked", Value: fdTotalPrincipal },
      { Metric: "FD Expected at Maturity", Value: fdTotalMaturity },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Summary");

    XLSX.writeFile(wb, `campcash-export-${new Date().toISOString().slice(0, 10)}.xlsx`);
    setSettingsMessage("Excel workbook exported with Ledger + pivot-style analysis sheets.");
  }

  function triggerJSONRestore() {
    if (jsonFileInputRef.current) jsonFileInputRef.current.click();
  }

  function handleJSONFileSelected(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result));
        if (!Array.isArray(parsed.transactions)) {
          setSettingsMessage("That doesn't look like a CampCash backup file.");
          return;
        }
        setRestoreSource("file");
        setPendingRestore(parsed);
        setRestoreConfirming(true);
      } catch (err) {
        setSettingsMessage("Couldn't read that JSON file.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  function confirmRestore() {
    const saved = pendingRestore;
    if (!saved) return;
    if (Array.isArray(saved.transactions)) {
      setTransactions(saved.transactions);
      const maxId = saved.transactions.length ? Math.max(...saved.transactions.map(t => t.id || 0)) : 0;
      setNextId(typeof saved.nextId === "number" ? saved.nextId : maxId + 1);
    }
    if (Array.isArray(saved.sources)) setSources(saved.sources);
    if (Array.isArray(saved.categories)) setCategories(saved.categories);
    if (saved.currency) setCurrency(saved.currency);
    if (typeof saved.openingBalance === "number") {
      setOpeningBalance(saved.openingBalance);
      setOpeningBalanceInput(String(saved.openingBalance));
    }
    if (saved.account) {
      if (typeof saved.account.name === "string") setAccountName(saved.account.name);
      if (typeof saved.account.email === "string") setAccountEmail(saved.account.email);
    }
    if (typeof saved.scenario === "string") setScenario(saved.scenario);
    if (typeof saved.lastRolloverMonth === "string") setLastRolloverMonth(saved.lastRolloverMonth);
    if (Array.isArray(saved.fds)) setFds(saved.fds);
    if (typeof saved.nextFdId === "number") setNextFdId(saved.nextFdId);
    if (restoreSource === "cloud") setCloudLastSyncedAt(saved.savedAt || new Date().toISOString());
    setPendingRestore(null);
    setRestoreConfirming(false);
    setSettingsMessage(restoreSource === "cloud" ? "Cloud data loaded." : "Backup restored.");
  }

  function cancelRestore() {
    setPendingRestore(null);
    setRestoreConfirming(false);
  }

  function triggerCSVImport() {
    if (csvFileInputRef.current) csvFileInputRef.current.click();
  }

  function handleCSVFileSelected(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const rows = parseCSV(String(reader.result));
        if (rows.length < 2) {
          setSettingsMessage("That CSV doesn't have any data rows.");
          return;
        }
        const headers = rows[0].map(h => h.trim().toLowerCase());
        const findCol = names => headers.findIndex(h => names.includes(h));
        const dateIdx = findCol(["date"]);
        const typeIdx = findCol(["type"]);
        const categoryIdx = findCol(["category", "header"]);
        const amountIdx = findCol(["amount"]);
        const notesIdx = findCol(["notes", "note"]);
        const methodIdx = findCol(["paymentmethod", "payment method", "method"]);

        if (amountIdx === -1 || typeIdx === -1) {
          setSettingsMessage("CSV needs at least a Type and an Amount column.");
          return;
        }

        let imported = 0, skipped = 0, idCounter = nextId;
        const newTx = [];
        rows.slice(1).forEach(r => {
          const rawType = (r[typeIdx] || "").trim().toLowerCase();
          const type = ["income", "in", "credit"].includes(rawType) ? "income"
            : ["expense", "out", "debit"].includes(rawType) ? "expense" : null;
          const amount = parseFloat(r[amountIdx]);
          if (!type || !amount || amount <= 0) { skipped++; return; }

          let isoDate = new Date().toISOString().slice(0, 10);
          if (dateIdx !== -1 && r[dateIdx]) {
            const d = new Date(r[dateIdx].trim());
            if (!isNaN(d.getTime())) isoDate = d.toISOString().slice(0, 10);
          }

          newTx.push({
            id: idCounter++,
            date: isoDate,
            type,
            amount,
            category: categoryIdx !== -1 && r[categoryIdx] ? r[categoryIdx].trim() : "Other",
            notes: notesIdx !== -1 ? (r[notesIdx] || "").trim() : "",
            ...(type === "expense" ? { paymentMethod: (methodIdx !== -1 && r[methodIdx] ? r[methodIdx].trim() : "Other") } : {}),
          });
          imported++;
        });

        if (newTx.length) {
          setTransactions(prev => [...prev, ...newTx]);
          setNextId(idCounter);
        }
        setSettingsMessage(
          `Imported ${imported} transaction${imported === 1 ? "" : "s"}` +
          (skipped ? `, skipped ${skipped} row${skipped === 1 ? "" : "s"} that couldn't be read.` : ".")
        );
      } catch (err) {
        setSettingsMessage("Couldn't read that CSV file.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  function exportPDFSummary() {
    try {
      const rows = displayList.map(t => `
        <tr>
          <td>${formatDate(t.date)}</td>
          <td>${t.type === "income" ? "Income" : "Expense"}</td>
          <td>${t.category}</td>
          <td>${t.paymentMethod || ""}</td>
          <td style="text-align:right">${t.type === "income" ? "+" : "-"} ${formatMoney(t.amount)}</td>
          <td style="text-align:right">${formatMoney(t.balanceAfter)}</td>
          <td>${t.notes || ""}</td>
        </tr>`).join("");
      const html = `
        <html>
          <head>
            <title>CampCash Summary</title>
            <style>
              body { font-family: -apple-system, Segoe UI, sans-serif; padding: 32px; color: #171717; }
              h1 { font-size: 20px; margin-bottom: 4px; }
              .muted { color: #737373; font-size: 12px; margin-bottom: 24px; }
              .stats { display: flex; gap: 24px; margin-bottom: 24px; }
              .stat { border: 1px solid #e5e5e5; border-radius: 8px; padding: 12px 16px; }
              .stat .label { font-size: 11px; color: #737373; }
              .stat .value { font-size: 18px; font-weight: 700; }
              table { width: 100%; border-collapse: collapse; font-size: 12px; }
              th, td { border-bottom: 1px solid #e5e5e5; padding: 6px 8px; text-align: left; }
              th { color: #737373; font-weight: 600; }
              @media print { body { padding: 12px; } }
            </style>
          </head>
          <body>
            <h1>CampCash — Account Summary${accountName ? " for " + accountName : ""}</h1>
            <div class="muted">Generated ${new Date().toLocaleString(currency.locale)}</div>
            <div class="stats">
              <div class="stat"><div class="label">Current Balance</div><div class="value">${formatMoney(balance)}</div></div>
              <div class="stat"><div class="label">Money In</div><div class="value">${formatMoney(periodIn)}</div></div>
              <div class="stat"><div class="label">Money Out</div><div class="value">${formatMoney(periodOut)}</div></div>
            </div>
            <table>
              <thead><tr><th>Date</th><th>Type</th><th>Header</th><th>Method</th><th style="text-align:right">Amount</th><th style="text-align:right">Balance After</th><th>Notes</th></tr></thead>
              <tbody>${rows || '<tr><td colspan="7" style="text-align:center;color:#737373">No transactions logged yet.</td></tr>'}</tbody>
            </table>
          </body>
        </html>
      `;

      // A hidden same-page iframe rather than window.open — popup blockers can't touch this
      // since no new browsing context is opened, just a normal DOM element.
      const iframe = document.createElement("iframe");
      iframe.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden;";
      document.body.appendChild(iframe);
      const doc = iframe.contentDocument || iframe.contentWindow.document;
      doc.open();
      doc.write(html);
      doc.close();

      setTimeout(() => {
        try {
          iframe.contentWindow.focus();
          iframe.contentWindow.print();
        } catch (err) {
          setSettingsMessage("Couldn't open the print dialog in this browser.");
        }
        setTimeout(() => {
          if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
        }, 1000);
      }, 300);

      setSettingsMessage("Opened the print dialog — choose \"Save as PDF\" as the destination.");
    } catch (e) {
      setSettingsMessage("Couldn't generate the PDF summary in this browser.");
    }
  }

  const bg = dark ? "bg-neutral-950" : "bg-white";
  const cardBg = dark ? "bg-neutral-900" : "bg-white";
  const border = dark ? "border-neutral-800" : "border-neutral-200";
  const text = dark ? "text-neutral-100" : "text-neutral-900";
  const textMuted = dark ? "text-neutral-400" : "text-neutral-500";
  const sidebarBg = dark ? "bg-neutral-950" : "bg-white";
  const inputBg = dark ? "bg-neutral-800 border-neutral-700 text-neutral-100 placeholder-neutral-500" : "bg-white border-neutral-200 text-neutral-900 placeholder-neutral-400";
  const rowHover = dark ? "hover:bg-neutral-800/60" : "hover:bg-neutral-50";
  const trackBg = dark ? "bg-neutral-800" : "bg-neutral-100";

  const navItems = [
    { key: "home", icon: Home, label: "Home" },
    { key: "fds", icon: Landmark, label: "Fixed Deposits", mobileLabel: "FDs" },
    { key: "budget", icon: PieChartIcon, label: "Budget Plan", mobileLabel: "Budget" },
    { key: "analytics", icon: BarChart2, label: "Analytics" },
    { key: "settings", icon: Settings, label: "Settings" },
  ];

  return (
    <div className={`min-h-screen w-full ${bg} ${text} font-sans flex`} style={{ fontFamily: "'Inter', 'Segoe UI', system-ui, sans-serif" }}>
      {/* Sidebar */}
      <aside className={`hidden md:flex w-56 shrink-0 border-r ${border} ${sidebarBg} flex-col justify-between py-6 px-4`}>
        <div>
          <div className="flex items-center gap-2 px-2 mb-8">
            <div className="w-8 h-8 rounded-lg bg-neutral-900 flex items-center justify-center text-white">
              <Wallet size={16} strokeWidth={2.2} />
            </div>
            <span className="font-semibold text-[15px] tracking-tight">CampCash</span>
          </div>
          <nav className="flex flex-col gap-1">
            {navItems.map(({ icon: Icon, label, key }) => (
              <button
                key={key}
                onClick={() => setActiveView(key)}
                className={`flex items-center gap-3 px-3 py-2.5 rounded-lg text-[14px] font-medium transition-colors text-left
                  ${activeView === key
                    ? (dark ? "bg-blue-500/10 text-blue-400" : "bg-blue-50 text-blue-600")
                    : `${textMuted} hover:${dark ? "bg-neutral-900" : "bg-neutral-50"}`}`}
              >
                <Icon size={18} strokeWidth={2} />
                {label}
              </button>
            ))}
          </nav>
        </div>

        <div className="flex flex-col gap-3">
          <div className={`flex items-center gap-2 px-3 text-[11px] ${textMuted}`}>
            <span className={`w-1.5 h-1.5 rounded-full ${
              saveStatus === "saving" ? "bg-amber-400" : saveStatus === "error" ? "bg-red-500" : saveStatus === "saved" ? "bg-green-500" : "bg-neutral-300"
            }`} />
            {saveStatus === "saving" ? "Saving…" : saveStatus === "error" ? "Save failed — retrying" : saveStatus === "saved" ? "All changes saved" : "Loading…"}
          </div>
          <div className="relative">
            <button
              onClick={() => setPeriodOpen(o => !o)}
              className={`w-full rounded-lg border ${border} px-3 py-2.5 flex items-center gap-2 text-left hover:${dark ? "bg-neutral-900" : "bg-neutral-50"}`}
            >
              <Calendar size={15} className={textMuted} />
              <div className="leading-tight">
                <div className={`text-[11px] ${textMuted}`}>{periodMode === "lifetime" ? "Viewing" : "Viewing Month"}</div>
                <div className="text-[13px] font-medium">{periodLabel}</div>
              </div>
            </button>
            {periodOpen && (
              <div className={`absolute left-0 bottom-full mb-2 z-20 w-64 rounded-lg border ${border} ${cardBg} shadow-lg p-3`}>
                <div className="flex items-center gap-2 mb-3">
                  <button
                    onClick={() => setPeriodMode("month")}
                    className={`flex-1 px-3 py-1.5 rounded-lg text-[12.5px] font-medium border transition-colors ${
                      periodMode === "month" ? "bg-blue-600 border-blue-600 text-white" : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                    }`}
                  >
                    Month
                  </button>
                  <button
                    onClick={() => setPeriodMode("lifetime")}
                    className={`flex-1 px-3 py-1.5 rounded-lg text-[12.5px] font-medium border transition-colors ${
                      periodMode === "lifetime" ? "bg-blue-600 border-blue-600 text-white" : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                    }`}
                  >
                    Lifetime
                  </button>
                </div>
                {periodMode === "month" && (
                  <input
                    type="month"
                    value={periodMonthKey}
                    onChange={e => e.target.value && setPeriodMonthKey(e.target.value)}
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[13px] outline-none`}
                  />
                )}
                <div className={`text-[11px] ${textMuted} mt-2`}>
                  {periodMode === "lifetime"
                    ? "Money In/Out, Analytics, and the ledger show everything ever logged."
                    : "Money In/Out, Analytics, and the ledger scope to this month."}
                </div>
              </div>
            )}
          </div>
          {periodOpen && (
            <div className="fixed inset-0 z-10" onClick={() => setPeriodOpen(false)} />
          )}
          <button
            onClick={() => setDark(d => !d)}
            className={`flex items-center justify-between px-3 py-2.5 rounded-lg border ${border} text-[13px] font-medium`}
          >
            <span className="flex items-center gap-2">
              {dark ? <Moon size={15} /> : <Sun size={15} />}
              Dark Mode
            </span>
            <span className={`w-8 rounded-full relative transition-colors ${dark ? "bg-blue-500" : "bg-neutral-300"}`} style={{ height: 18 }}>
              <span className={`absolute top-0.5 w-3.5 h-3.5 rounded-full bg-white transition-all ${dark ? "left-4" : "left-0.5"}`} />
            </span>
          </button>
        </div>
      </aside>

      {/* Mobile bottom nav */}
      <nav className={`md:hidden fixed bottom-0 inset-x-0 z-20 border-t ${border} ${sidebarBg} flex items-stretch justify-around px-1`}>
        {navItems.map(({ icon: Icon, label, mobileLabel, key }) => (
          <button
            key={key}
            onClick={() => setActiveView(key)}
            className={`flex-1 flex flex-col items-center gap-0.5 py-2.5 text-[10px] font-medium ${
              activeView === key ? "text-blue-500" : textMuted
            }`}
          >
            <Icon size={18} strokeWidth={2} />
            {mobileLabel || label}
          </button>
        ))}
      </nav>

      {/* Main content */}
      <main className="flex-1 max-w-5xl mx-auto px-4 sm:px-8 py-6 pb-24 md:pb-6 w-full">
        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-[16px] truncate max-w-[55vw] sm:max-w-none">
              {activeView === "home" ? (accountName ? `${accountName}'s Wallet` : "My Wallet") : activeView === "fds" ? "Fixed Deposits" : activeView === "budget" ? "5-Year Budget Plan" : activeView === "analytics" ? "Analytics" : "Settings"}
            </span>
            {activeView === "home" && <span className="w-1.5 h-1.5 rounded-full bg-green-500" />}
          </div>
          <div className="flex items-center gap-4 relative">
            <div className="relative">
              <button onClick={toggleNotifPanel} className="relative">
                <Bell size={18} className={textMuted} />
                {unseenNotifCount > 0 && (
                  <span className="absolute -top-1 -right-1 w-4 h-4 rounded-full bg-red-500 text-white text-[9px] font-bold flex items-center justify-center">
                    {unseenNotifCount}
                  </span>
                )}
              </button>
              {notifOpen && (
                <div className={`absolute right-0 top-8 z-20 w-80 max-w-[85vw] max-h-80 overflow-y-auto rounded-lg border ${border} ${cardBg} shadow-lg py-2`}>
                  <div className={`px-3 pb-2 text-[12px] font-semibold ${textMuted}`}>Notifications</div>
                  {notifications.length === 0 ? (
                    <div className={`px-3 py-4 text-[13px] ${textMuted} text-center`}>You're all caught up.</div>
                  ) : (
                    notifications.map(n => (
                      <div key={n.id} className={`px-3 py-2 text-[12.5px] flex items-start gap-2 ${rowHover}`}>
                        <span className={`mt-1 w-1.5 h-1.5 rounded-full shrink-0 ${n.tone === "warning" ? "bg-red-500" : "bg-blue-500"}`} />
                        <span>{n.text}</span>
                      </div>
                    ))
                  )}
                </div>
              )}
            </div>

            <div className="relative">
              <button onClick={() => { setMenuOpen(o => !o); setNotifOpen(false); }}>
                <MoreVertical size={18} className={textMuted} />
              </button>
              {menuOpen && (
                <div className={`absolute right-0 top-8 z-20 w-56 max-w-[85vw] rounded-lg border ${border} ${cardBg} shadow-lg py-1`}>
                  <button
                    onClick={() => { setDark(d => !d); setMenuOpen(false); }}
                    className={`w-full flex items-center gap-2 text-left px-3 py-2 text-[13px] ${rowHover}`}
                  >
                    {dark ? <Sun size={14} /> : <Moon size={14} />} {dark ? "Switch to Light Mode" : "Switch to Dark Mode"}
                  </button>
                  <button
                    onClick={() => { setActiveView("settings"); setMenuOpen(false); }}
                    className={`w-full flex items-center gap-2 text-left px-3 py-2 text-[13px] ${rowHover}`}
                  >
                    <Settings size={14} /> Go to Settings
                  </button>
                  <button
                    onClick={() => { exportPDFSummary(); setMenuOpen(false); }}
                    className={`w-full flex items-center gap-2 text-left px-3 py-2 text-[13px] ${rowHover}`}
                  >
                    <FileText size={14} /> Export PDF Summary
                  </button>
                  <button
                    onClick={() => { exportJSONBackup(); setMenuOpen(false); }}
                    className={`w-full flex items-center gap-2 text-left px-3 py-2 text-[13px] ${rowHover}`}
                  >
                    <Download size={14} /> Download JSON Backup
                  </button>
                  <div className={`my-1 border-t ${border}`} />
                  <button
                    onClick={() => { setActiveView("settings"); setResetConfirming(true); setMenuOpen(false); }}
                    className={`w-full flex items-center gap-2 text-left px-3 py-2 text-[13px] text-red-500 ${rowHover}`}
                  >
                    <AlertTriangle size={14} /> Reset All Data
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>

        {(notifOpen || menuOpen) && (
          <div className="fixed inset-0 z-10" onClick={() => { setNotifOpen(false); setMenuOpen(false); }} />
        )}

        {/* Mobile time-period selector — the sidebar version is hidden below md, so it needs its own trigger here */}
        <div className="md:hidden relative mb-4">
          <button
            onClick={() => setPeriodOpen(o => !o)}
            className={`w-full rounded-lg border ${border} ${cardBg} px-3 py-2.5 flex items-center gap-2 text-left`}
          >
            <Calendar size={15} className={textMuted} />
            <div className="leading-tight">
              <div className={`text-[11px] ${textMuted}`}>{periodMode === "lifetime" ? "Viewing" : "Viewing Month"}</div>
              <div className="text-[13px] font-medium">{periodLabel}</div>
            </div>
          </button>
          {periodOpen && (
            <div className={`absolute left-0 top-full mt-2 z-20 w-64 max-w-[85vw] rounded-lg border ${border} ${cardBg} shadow-lg p-3`}>
              <div className="flex items-center gap-2 mb-3">
                <button
                  onClick={() => setPeriodMode("month")}
                  className={`flex-1 px-3 py-1.5 rounded-lg text-[12.5px] font-medium border transition-colors ${
                    periodMode === "month" ? "bg-blue-600 border-blue-600 text-white" : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                  }`}
                >
                  Month
                </button>
                <button
                  onClick={() => setPeriodMode("lifetime")}
                  className={`flex-1 px-3 py-1.5 rounded-lg text-[12.5px] font-medium border transition-colors ${
                    periodMode === "lifetime" ? "bg-blue-600 border-blue-600 text-white" : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                  }`}
                >
                  Lifetime
                </button>
              </div>
              {periodMode === "month" && (
                <input
                  type="month"
                  value={periodMonthKey}
                  onChange={e => e.target.value && setPeriodMonthKey(e.target.value)}
                  className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[13px] outline-none`}
                />
              )}
              <div className={`text-[11px] ${textMuted} mt-2`}>
                {periodMode === "lifetime"
                  ? "Money In/Out, Analytics, and the ledger show everything ever logged."
                  : "Money In/Out, Analytics, and the ledger scope to this month."}
              </div>
            </div>
          )}
          {periodOpen && (
            <div className="fixed inset-0 z-10" onClick={() => setPeriodOpen(false)} />
          )}
        </div>

        {activeView === "home" && (
          <>
            {/* Balance hero */}
            <div className="rounded-2xl bg-gradient-to-br from-blue-500 to-blue-600 px-5 sm:px-7 py-6 flex items-center justify-between mb-5 shadow-sm">
              <div>
                <div className="text-blue-100 text-[12px] font-medium tracking-wide uppercase mb-2">Current Balance</div>
                <div className="text-white text-3xl sm:text-4xl font-bold tracking-tight mb-3">{formatMoney(balance)}</div>
                <div className="inline-flex items-center gap-1.5 bg-white/15 text-blue-50 text-[12px] font-medium px-3 py-1 rounded-full">
                  <span className="w-1.5 h-1.5 rounded-full bg-green-300" />
                  Available to spend
                </div>
              </div>
              <div className="hidden sm:flex w-20 h-20 rounded-2xl bg-white/15 items-center justify-center">
                <Wallet size={36} className="text-white/90" strokeWidth={1.6} />
              </div>
            </div>

            {/* Stats row */}
            <div className={`rounded-2xl border ${border} ${cardBg} px-6 py-5 mb-5 grid grid-cols-1 md:grid-cols-3 divide-y md:divide-y-0 md:divide-x ${dark ? "divide-neutral-800" : "divide-neutral-200"}`}>
              <div className="flex items-center gap-3 pb-4 md:pb-0 md:pr-4">
                <div className={`w-10 h-10 rounded-full flex items-center justify-center ${dark ? "bg-green-500/10" : "bg-green-50"}`}>
                  <ArrowDownCircle size={18} className="text-green-500" />
                </div>
                <div>
                  <div className="text-[13px] font-semibold text-green-500">Money In</div>
                  <div className="text-lg font-bold">{formatMoney(periodIn)}</div>
                  <div className={`text-[11px] ${textMuted}`}>{periodLabel}</div>
                </div>
              </div>
              <div className="flex items-center gap-3 py-4 md:py-0 md:px-4">
                <div className={`w-10 h-10 rounded-full flex items-center justify-center ${dark ? "bg-red-500/10" : "bg-red-50"}`}>
                  <ArrowUpCircle size={18} className="text-red-500" />
                </div>
                <div>
                  <div className="text-[13px] font-semibold text-red-500">Money Out</div>
                  <div className="text-lg font-bold">{formatMoney(periodOut)}</div>
                  <div className={`text-[11px] ${textMuted}`}>{periodLabel}</div>
                </div>
              </div>
              <div className="flex items-center gap-3 pt-4 md:pt-0 md:pl-4">
                <div className={`w-10 h-10 rounded-full flex items-center justify-center ${dark ? "bg-blue-500/10" : "bg-blue-50"}`}>
                  <Wallet size={18} className="text-blue-500" />
                </div>
                <div>
                  <div className="text-[13px] font-semibold text-blue-500">Net</div>
                  <div className="text-lg font-bold">{formatMoney(periodIn - periodOut)}</div>
                  <div className={`text-[11px] ${textMuted}`}>{periodLabel}</div>
                </div>
              </div>
            </div>

            {error && (
              <div className="mb-5 rounded-lg border border-red-300 bg-red-50 text-red-700 text-[13px] px-4 py-2.5">
                {error}
              </div>
            )}

            {/* Add / Spend forms */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5 mb-5">
              <div className={`rounded-2xl border ${dark ? "border-green-900 bg-green-500/5" : "border-green-100 bg-green-50/50"} p-5`}>
                <div className="flex items-center gap-3 mb-4">
                  <div className={`w-9 h-9 rounded-full flex items-center justify-center ${dark ? "bg-green-500/15" : "bg-green-100"}`}>
                    <ArrowDownCircle size={17} className="text-green-600" />
                  </div>
                  <div>
                    <div className="font-semibold text-[14px] text-green-700">Add Money</div>
                    <div className={`text-[12px] ${textMuted}`}>Add funds to your wallet</div>
                  </div>
                </div>
                <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Amount</label>
                <div className={`flex items-center rounded-lg border ${inputBg} px-3 mb-3`}>
                  <span className={textMuted}>{currency.symbol}</span>
                  <input
                    type="number"
                    value={addAmount}
                    onChange={e => setAddAmount(e.target.value)}
                    placeholder="0.00"
                    className="w-full bg-transparent outline-none px-2 py-2 text-[14px]"
                  />
                </div>
                <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Source</label>
                <select
                  value={addSource}
                  onChange={e => setAddSource(e.target.value)}
                  className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] mb-3 outline-none`}
                >
                  <option value="">Select source</option>
                  {sources.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
                <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Notes (optional)</label>
                <input
                  type="text"
                  value={addNotes}
                  onChange={e => setAddNotes(e.target.value)}
                  placeholder="e.g. Parents, Scholarship, Refund"
                  className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] mb-4 outline-none`}
                />
                <button
                  onClick={handleDeposit}
                  className="w-full bg-green-600 hover:bg-green-700 text-white font-semibold text-[14px] py-2.5 rounded-lg flex items-center justify-center gap-2 transition-colors"
                >
                  <Plus size={16} /> Deposit
                </button>
              </div>

              <div className={`rounded-2xl border ${dark ? "border-red-900 bg-red-500/5" : "border-red-100 bg-red-50/50"} p-5`}>
                <div className="flex items-center gap-3 mb-4">
                  <div className={`w-9 h-9 rounded-full flex items-center justify-center ${dark ? "bg-red-500/15" : "bg-red-100"}`}>
                    <ArrowUpCircle size={17} className="text-red-600" />
                  </div>
                  <div>
                    <div className="font-semibold text-[14px] text-red-700">Spend Money</div>
                    <div className={`text-[12px] ${textMuted}`}>Deduct money from your wallet</div>
                  </div>
                </div>
                <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Amount</label>
                <div className={`flex items-center rounded-lg border ${inputBg} px-3 mb-3`}>
                  <span className={textMuted}>{currency.symbol}</span>
                  <input
                    type="number"
                    value={spendAmount}
                    onChange={e => setSpendAmount(e.target.value)}
                    placeholder="0.00"
                    className="w-full bg-transparent outline-none px-2 py-2 text-[14px]"
                  />
                </div>
                <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Category</label>
                <select
                  value={spendCategory}
                  onChange={e => setSpendCategory(e.target.value)}
                  className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] mb-3 outline-none`}
                >
                  <option value="">Select category</option>
                  {categories.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
                <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Payment Method</label>
                <div className="flex items-center gap-2 mb-3">
                  {PAYMENT_METHODS.map(m => (
                    <button
                      key={m}
                      type="button"
                      onClick={() => setSpendMethod(m)}
                      className={`flex-1 px-2 py-1.5 rounded-lg text-[12.5px] font-medium border transition-colors ${
                        spendMethod === m
                          ? "bg-red-500 border-red-500 text-white"
                          : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                      }`}
                    >
                      {m}
                    </button>
                  ))}
                </div>
                <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Notes (optional)</label>
                <input
                  type="text"
                  value={spendNotes}
                  onChange={e => setSpendNotes(e.target.value)}
                  placeholder="e.g. Lunch, Transport, Stationery"
                  className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] mb-4 outline-none`}
                />
                <button
                  onClick={handleSpend}
                  className="w-full bg-red-500 hover:bg-red-600 text-white font-semibold text-[14px] py-2.5 rounded-lg flex items-center justify-center gap-2 transition-colors"
                >
                  <Minus size={16} /> Spend
                </button>
              </div>
            </div>

            {/* Quick add */}
            <div className={`rounded-2xl border ${border} ${cardBg} px-6 py-4 mb-5 flex items-center justify-between flex-wrap gap-3`}>
              <div className="flex items-center gap-3 flex-wrap">
                <span className="font-semibold text-[14px]">Quick Add</span>
                {[100, 500, 1000].map(a => (
                  <button
                    key={"in" + a}
                    onClick={() => handleQuickAdd("income", a)}
                    className={`text-[13px] font-medium px-3.5 py-1.5 rounded-full ${dark ? "bg-green-500/10 text-green-400 hover:bg-green-500/20" : "bg-green-50 text-green-700 hover:bg-green-100"}`}
                  >
                    + {currency.symbol}{a.toLocaleString(currency.locale)}
                  </button>
                ))}
                <span className={`w-px h-5 ${dark ? "bg-neutral-800" : "bg-neutral-200"}`} />
                {[50, 100, 200].map(a => (
                  <button
                    key={"out" + a}
                    onClick={() => handleQuickAdd("expense", a)}
                    className={`text-[13px] font-medium px-3.5 py-1.5 rounded-full ${dark ? "bg-red-500/10 text-red-400 hover:bg-red-500/20" : "bg-red-50 text-red-700 hover:bg-red-100"}`}
                  >
                    − {currency.symbol}{a}
                  </button>
                ))}
              </div>
              <span className={`text-[12px] ${textMuted}`}>Tap amount to add transaction</span>
            </div>

            {/* Ledger */}
            <div className={`rounded-2xl border ${border} ${cardBg} overflow-hidden`}>
              <div className="flex items-center justify-between px-6 py-4 flex-wrap gap-3">
                <span className="font-semibold text-[15px]">Transaction Ledger</span>
                <div className="flex items-center gap-2 relative">
                  <div className={`flex items-center gap-2 rounded-lg border ${inputBg} px-3 py-1.5`}>
                    <Search size={14} className={textMuted} />
                    <input
                      value={search}
                      onChange={e => setSearch(e.target.value)}
                      placeholder="Search notes..."
                      className="bg-transparent outline-none text-[13px] w-36"
                    />
                  </div>
                  <button
                    onClick={() => setFilterOpen(o => !o)}
                    className={`rounded-lg border p-2 ${filterCategory !== "All" ? "border-blue-500 text-blue-500" : `${border} ${textMuted}`}`}
                  >
                    <SlidersHorizontal size={14} />
                  </button>
                  {filterOpen && (
                    <div className={`absolute right-0 top-11 z-10 w-64 max-w-[85vw] max-h-72 overflow-y-auto rounded-lg border ${border} ${cardBg} shadow-lg py-1`}>
                      <button
                        onClick={() => { setFilterCategory("All"); setFilterOpen(false); }}
                        className={`w-full text-left px-3 py-2 text-[13px] ${rowHover} ${filterCategory === "All" ? "text-blue-500 font-medium" : ""}`}
                      >
                        All headers
                      </button>
                      <div className={`px-3 pt-2 pb-1 text-[11px] font-medium ${textMuted}`}>Income sources</div>
                      {sources.map(h => (
                        <button
                          key={h}
                          onClick={() => { setFilterCategory(h); setFilterOpen(false); }}
                          className={`w-full text-left px-3 py-2 text-[13px] ${rowHover} ${filterCategory === h ? "text-blue-500 font-medium" : ""}`}
                        >
                          {h}
                        </button>
                      ))}
                      <div className={`px-3 pt-2 pb-1 text-[11px] font-medium ${textMuted}`}>Expense categories</div>
                      {categories.map(h => (
                        <button
                          key={h}
                          onClick={() => { setFilterCategory(h); setFilterOpen(false); }}
                          className={`w-full text-left px-3 py-2 text-[13px] ${rowHover} ${filterCategory === h ? "text-blue-500 font-medium" : ""}`}
                        >
                          {h}
                        </button>
                      ))}
                      <div className={`px-3 pt-2 pb-1 text-[11px] font-medium ${textMuted}`}>System</div>
                      <button
                        onClick={() => { setFilterCategory(ROLLOVER_CATEGORY); setFilterOpen(false); }}
                        className={`w-full text-left px-3 py-2 text-[13px] ${rowHover} ${filterCategory === ROLLOVER_CATEGORY ? "text-blue-500 font-medium" : ""}`}
                      >
                        {ROLLOVER_CATEGORY}
                      </button>
                      <button
                        onClick={() => { setFilterCategory(FD_CATEGORY); setFilterOpen(false); }}
                        className={`w-full text-left px-3 py-2 text-[13px] ${rowHover} ${filterCategory === FD_CATEGORY ? "text-blue-500 font-medium" : ""}`}
                      >
                        {FD_CATEGORY}
                      </button>
                    </div>
                  )}
                </div>
              </div>
              {filterCategory !== "All" && (
                <div className="px-6 pb-3 -mt-1 flex items-center gap-2">
                  <span className={`inline-flex items-center gap-1.5 text-[12px] font-medium px-2.5 py-1 rounded-full ${dark ? "bg-blue-500/10 text-blue-400" : "bg-blue-50 text-blue-600"}`}>
                    {filterCategory}
                    <button onClick={() => setFilterCategory("All")} className="hover:opacity-70">×</button>
                  </span>
                </div>
              )}
              <div className="overflow-x-auto">
              <table className="w-full text-[13.5px]">
                <thead>
                  <tr className={`border-t border-b ${border} ${textMuted}`}>
                    <th className="text-left font-medium px-6 py-2.5">
                      <button onClick={() => toggleSort("date")} className="flex items-center gap-1 hover:opacity-70">
                        Date {sortKey === "date" && (sortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                      </button>
                    </th>
                    <th className="text-left font-medium py-2.5">Type</th>
                    <th className="text-left font-medium py-2.5">
                      <button onClick={() => toggleSort("amount")} className="flex items-center gap-1 hover:opacity-70">
                        Amount {sortKey === "amount" && (sortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
                      </button>
                    </th>
                    <th className="text-left font-medium py-2.5">Balance After</th>
                    <th className="text-left font-medium py-2.5">Notes</th>
                    <th className="py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {displayList.length === 0 && (
                    <tr>
                      <td colSpan={6} className={`text-center py-8 text-[13px] ${textMuted}`}>
                        {transactions.length === 0 ? (
                          <div className="flex flex-col items-center gap-2">
                            <span>No transactions yet — add one above to get started.</span>
                            <button onClick={loadSampleData} className="text-blue-600 font-medium hover:text-blue-700">
                              Or load sample data to explore the app
                            </button>
                          </div>
                        ) : `No transactions match your current filters${periodMode === "month" ? ` for ${periodLabel}` : ""}.`}
                      </td>
                    </tr>
                  )}
                  {displayList.map(t => (
                    <tr key={t.id} className={`border-b ${border} ${rowHover} transition-colors`}>
                      <td className="px-6 py-3">{formatDate(t.date)}</td>
                      <td className="py-3">
                        <span className={`inline-flex items-center justify-center w-6 h-6 rounded-full ${t.type === "income" ? (dark ? "bg-green-500/15" : "bg-green-100") : (dark ? "bg-red-500/15" : "bg-red-100")}`}>
                          {t.type === "income"
                            ? <ArrowDownCircle size={13} className="text-green-600" />
                            : <ArrowUpCircle size={13} className="text-red-600" />}
                        </span>
                      </td>
                      <td className={`py-3 font-semibold ${t.type === "income" ? "text-green-600" : "text-red-600"}`}>
                        {t.type === "income" ? "+ " : "- "}{formatMoney(t.amount)}
                      </td>
                      <td className="py-3">{formatMoney(t.balanceAfter)}</td>
                      <td className={`py-3 ${textMuted}`}>
                        {t.notes || t.category}
                        {t.paymentMethod && (
                          <span className={`ml-2 text-[10px] px-1.5 py-0.5 rounded ${dark ? "bg-neutral-800 text-neutral-400" : "bg-neutral-100 text-neutral-500"}`}>
                            {t.paymentMethod}
                          </span>
                        )}
                      </td>
                      <td className="py-3 pr-4 text-right">
                        <button onClick={() => deleteTransaction(t.id)} className={`${textMuted} hover:text-red-500`}>
                          <Trash2 size={14} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              </div>
              <div className="text-center py-4">
                <button className="text-blue-600 text-[13px] font-medium inline-flex items-center gap-1">
                  View All Transactions <ChevronDown size={14} />
                </button>
              </div>
            </div>
          </>
        )}

        {activeView === "fds" && (
          <>
            {/* Summary cards */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-5 mb-5">
              <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                <div className={`text-[12px] ${textMuted} mb-1`}>Active FDs</div>
                <div className="text-2xl font-bold">{activeFds.length}</div>
              </div>
              <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                <div className={`text-[12px] ${textMuted} mb-1`}>Total Principal Locked</div>
                <div className="text-2xl font-bold">{formatMoney(fdTotalPrincipal)}</div>
              </div>
              <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                <div className={`text-[12px] ${textMuted} mb-1`}>Expected at Maturity</div>
                <div className="text-2xl font-bold">{formatMoney(fdTotalMaturity)}</div>
              </div>
              <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                <div className={`text-[12px] ${textMuted} mb-1`}>Expected Interest</div>
                <div className="text-2xl font-bold text-green-600">{formatMoney(fdTotalInterest)}</div>
              </div>
            </div>

            {fdError && (
              <div className="mb-5 rounded-lg border border-red-300 bg-red-50 text-red-700 text-[13px] px-4 py-2.5">
                {fdError}
              </div>
            )}

            {/* Add FD form */}
            <div className={`rounded-2xl border ${dark ? "border-blue-900 bg-blue-500/5" : "border-blue-100 bg-blue-50/50"} p-5 mb-5`}>
              <div className="flex items-center gap-3 mb-4">
                <div className={`w-9 h-9 rounded-full flex items-center justify-center ${dark ? "bg-blue-500/15" : "bg-blue-100"}`}>
                  <Landmark size={17} className="text-blue-600" />
                </div>
                <div>
                  <div className="font-semibold text-[14px] text-blue-700">Add Fixed Deposit</div>
                  <div className={`text-[12px] ${textMuted}`}>Track a new FD and optionally debit it from your wallet</div>
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-3">
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Bank / Label</label>
                  <input
                    type="text"
                    value={fdLabel}
                    onChange={e => setFdLabel(e.target.value)}
                    placeholder="e.g. SBI 1-year FD"
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] outline-none`}
                  />
                </div>
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Principal</label>
                  <div className={`flex items-center rounded-lg border ${inputBg} px-3`}>
                    <span className={textMuted}>{currency.symbol}</span>
                    <input
                      type="number"
                      value={fdPrincipal}
                      onChange={e => setFdPrincipal(e.target.value)}
                      placeholder="0.00"
                      className="w-full bg-transparent outline-none px-2 py-2 text-[14px]"
                    />
                  </div>
                </div>
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Annual Interest Rate</label>
                  <div className={`flex items-center rounded-lg border ${inputBg} px-3`}>
                    <input
                      type="number"
                      value={fdRate}
                      onChange={e => setFdRate(e.target.value)}
                      placeholder="7.0"
                      step="0.05"
                      className="w-full bg-transparent outline-none px-2 py-2 text-[14px]"
                    />
                    <span className={textMuted}>%</span>
                  </div>
                </div>
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Tenure (months)</label>
                  <input
                    type="number"
                    value={fdTenure}
                    onChange={e => setFdTenure(e.target.value)}
                    placeholder="12"
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] outline-none`}
                  />
                </div>
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Start Date</label>
                  <input
                    type="date"
                    value={fdStartDate}
                    onChange={e => setFdStartDate(e.target.value)}
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] outline-none`}
                  />
                </div>
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Compounding</label>
                  <select
                    value={fdCompounding}
                    onChange={e => setFdCompounding(e.target.value)}
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] outline-none`}
                  >
                    <option value="monthly">Monthly</option>
                    <option value="quarterly">Quarterly</option>
                    <option value="annually">Annually</option>
                    <option value="simple">Simple Interest</option>
                  </select>
                </div>
              </div>
              <label className="flex items-center gap-2 text-[13px] mb-4 cursor-pointer">
                <input type="checkbox" checked={fdDebitOnCreate} onChange={e => setFdDebitOnCreate(e.target.checked)} className="w-4 h-4" />
                Debit this principal from my wallet now
              </label>
              <button
                onClick={addFD}
                className="px-5 py-2.5 rounded-lg text-[14px] font-semibold bg-blue-600 hover:bg-blue-700 text-white flex items-center gap-2"
              >
                <Plus size={16} /> Add Fixed Deposit
              </button>
            </div>

            {/* FD list */}
            {fdsWithComputed.length === 0 ? (
              <div className={`rounded-2xl border ${border} ${cardBg} p-10 text-center`}>
                <Landmark size={28} className={`mx-auto mb-3 ${textMuted}`} />
                <div className="font-semibold text-[15px] mb-1">No fixed deposits yet</div>
                <div className={`text-[13px] ${textMuted}`}>Add one above to start tracking it toward maturity.</div>
              </div>
            ) : (
              <div className="flex flex-col gap-4">
                {fdsWithComputed.map(fd => (
                  <div key={fd.id} className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                    <div className="flex items-start justify-between gap-3 mb-3 flex-wrap">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="font-semibold text-[14px]">{fd.label}</span>
                          <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded ${
                            fd.status === "active"
                              ? (fd.isOverdue ? (dark ? "bg-amber-500/15 text-amber-400" : "bg-amber-100 text-amber-700") : (dark ? "bg-blue-500/15 text-blue-400" : "bg-blue-100 text-blue-700"))
                              : fd.status === "matured" ? (dark ? "bg-green-500/15 text-green-400" : "bg-green-100 text-green-700")
                              : (dark ? "bg-neutral-800 text-neutral-400" : "bg-neutral-100 text-neutral-600")
                          }`}>
                            {fd.status === "active" ? (fd.isOverdue ? "Matured — not closed" : "Active") : fd.status === "matured" ? "Matured" : "Withdrawn"}
                          </span>
                        </div>
                        <div className={`text-[12px] ${textMuted} mt-0.5`}>
                          {formatMoney(fd.principal)} at {fd.rate}% ({fd.compounding}) · {formatDate(fd.startDate)} → {formatDate(fd.maturityDate)}
                        </div>
                      </div>
                      <button onClick={() => deleteFd(fd.id)} className={`${textMuted} hover:text-red-500`}>
                        <Trash2 size={14} />
                      </button>
                    </div>

                    {fd.status === "active" ? (
                      <>
                        <div className={`w-full h-2 rounded-full ${trackBg} overflow-hidden mb-2`}>
                          <div className="h-full rounded-full bg-blue-500" style={{ width: `${fd.progressPct}%` }} />
                        </div>
                        <div className="flex items-center justify-between text-[12px] mb-4">
                          <span className={textMuted}>{fd.progressPct}% of tenure elapsed</span>
                          <span>
                            Expected: <span className="font-semibold">{formatMoney(fd.maturityAmount)}</span>
                            <span className="text-green-600"> (+{formatMoney(fd.expectedInterest)})</span>
                          </span>
                        </div>
                        {closingFdId === fd.id ? (
                          <div className={`rounded-lg border ${border} p-3`}>
                            <div className={`text-[12px] font-medium mb-2`}>
                              {closingDraft.status === "matured" ? "Mark as matured" : "Withdraw early"}
                            </div>
                            <label className={`text-[11px] font-medium ${textMuted} mb-1 block`}>Payout amount</label>
                            <div className={`flex items-center rounded-lg border ${inputBg} px-3 mb-2`}>
                              <span className={textMuted}>{currency.symbol}</span>
                              <input
                                type="number"
                                value={closingDraft.amount}
                                onChange={e => setClosingDraft(d => ({ ...d, amount: e.target.value }))}
                                className="w-full bg-transparent outline-none px-2 py-1.5 text-[13px]"
                              />
                            </div>
                            <label className="flex items-center gap-2 text-[12.5px] mb-3 cursor-pointer">
                              <input
                                type="checkbox"
                                checked={closingDraft.credit}
                                onChange={e => setClosingDraft(d => ({ ...d, credit: e.target.checked }))}
                                className="w-4 h-4"
                              />
                              Credit this payout to my wallet
                            </label>
                            <div className="flex items-center gap-2">
                              <button onClick={confirmCloseFd} className="px-3 py-1.5 rounded-lg text-[13px] font-medium bg-blue-600 text-white hover:bg-blue-700">
                                Confirm
                              </button>
                              <button onClick={cancelClosingFd} className={`px-3 py-1.5 rounded-lg text-[13px] font-medium border ${border} ${textMuted}`}>
                                Cancel
                              </button>
                            </div>
                          </div>
                        ) : (
                          <div className="flex items-center gap-2 flex-wrap">
                            <button
                              onClick={() => startClosingFd(fd.id, "matured", fd.maturityAmount)}
                              className="px-3 py-1.5 rounded-lg text-[12.5px] font-medium bg-green-600 text-white hover:bg-green-700"
                            >
                              Mark Matured
                            </button>
                            <button
                              onClick={() => startClosingFd(fd.id, "withdrawn", fd.principal)}
                              className={`px-3 py-1.5 rounded-lg text-[12.5px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                            >
                              Withdraw Early
                            </button>
                          </div>
                        )}
                      </>
                    ) : (
                      <div className="text-[12.5px]">
                        <span className={textMuted}>Closed {formatDate(fd.closedDate)} — payout </span>
                        <span className="font-semibold">{formatMoney(fd.payoutAmount)}</span>
                        <span className={textMuted}>{fd.credited ? " · credited to wallet" : " · not credited to wallet"}</span>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </>
        )}

        {activeView === "budget" && (
          <>
            {/* Scenario selector */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5 mb-5`}>
              <div className="flex items-center justify-between mb-3">
                <span className="font-semibold text-[14px]">Scenario</span>
                <span className={`text-[12px] ${textMuted}`}>Applies to personal & travel variable costs only</span>
              </div>
              <div className="flex items-center gap-2">
                {["Conservative", "Comfortable", "High"].map(s => (
                  <button
                    key={s}
                    onClick={() => setScenario(s)}
                    className={`px-4 py-2 rounded-lg text-[13px] font-medium border transition-colors ${
                      scenario === s
                        ? "bg-blue-600 border-blue-600 text-white"
                        : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                    }`}
                  >
                    {s} · {SCENARIO_MULTIPLIERS[s]}x
                  </button>
                ))}
              </div>
            </div>

            {/* Summary cards */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-5 mb-5">
              <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                <div className={`text-[12px] ${textMuted} mb-1`}>Selected Scenario</div>
                <div className="text-2xl font-bold">{scenario}</div>
                <div className={`text-[11px] ${textMuted} mt-1`}>{SCENARIO_MULTIPLIERS[scenario]}x on variable costs</div>
              </div>
              <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                <div className={`text-[12px] ${textMuted} mb-1`}>5-Year Total</div>
                <div className="text-2xl font-bold">{formatMoney(fiveYearTotal)}</div>
                <div className={`text-[11px] ${textMuted} mt-1`}>All heads, Years 1–5</div>
              </div>
              <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                <div className={`text-[12px] ${textMuted} mb-1`}>Suggested Monthly Set-Aside</div>
                <div className="text-2xl font-bold">{formatMoney(monthlySetAside)}</div>
                <div className={`text-[11px] ${textMuted} mt-1`}>5-year total ÷ 60 months</div>
              </div>
            </div>

            {/* Year selector */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5 mb-5`}>
              <div className="flex items-center justify-between mb-4">
                <span className="font-semibold text-[14px]">Year-wise totals</span>
              </div>
              <div className="flex items-center gap-2 mb-4">
                {[1, 2, 3, 4, 5].map(y => (
                  <button
                    key={y}
                    onClick={() => setBudgetYear(y)}
                    className={`px-4 py-2 rounded-lg text-[13px] font-medium border transition-colors ${
                      budgetYear === y
                        ? "bg-blue-600 border-blue-600 text-white"
                        : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                    }`}
                  >
                    Year {y}
                  </button>
                ))}
              </div>
              <div className="flex items-baseline gap-2">
                <span className="text-3xl font-bold">{formatMoney(yearTotals[budgetYear - 1])}</span>
                <span className={`text-[13px] ${textMuted}`}>budgeted for Year {budgetYear} at {scenario}</span>
              </div>
            </div>

            {/* Group summary */}
            <div className={`rounded-2xl border ${border} ${cardBg} overflow-hidden mb-5`}>
              <div className="px-6 py-4">
                <span className="font-semibold text-[15px]">Spending Groups — Year {budgetYear}</span>
              </div>
              <div className="overflow-x-auto">
              <table className="w-full text-[13.5px]">
                <thead>
                  <tr className={`border-t border-b ${border} ${textMuted}`}>
                    <th className="text-left font-medium px-6 py-2.5">Group</th>
                    <th className="text-left font-medium py-2.5">Year {budgetYear}</th>
                    <th className="text-left font-medium py-2.5">5-Year Total</th>
                  </tr>
                </thead>
                <tbody>
                  {groupTotals.map(g => (
                    <tr key={g.group} className={`border-b ${border} ${rowHover}`}>
                      <td className="px-6 py-3 font-medium">{g.group}</td>
                      <td className="py-3">{formatMoney(g.years[budgetYear - 1])}</td>
                      <td className={`py-3 ${textMuted}`}>{formatMoney(g.years.reduce((a, b) => a + b, 0))}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              </div>
            </div>

            {/* Budgeted vs Logged */}
            <div className={`rounded-2xl border ${border} ${cardBg} overflow-hidden mb-5`}>
              <div className="px-6 py-4">
                <span className="font-semibold text-[15px]">Budgeted vs Logged — Year {budgetYear}</span>
                <div className={`text-[12px] ${textMuted} mt-0.5`}>Comparing the {scenario} plan against what you've actually logged in the ledger so far</div>
              </div>
              <div className="px-6 pb-5 flex flex-col gap-3">
                {categories.filter(c => c !== "Transport / Other").map(cat => {
                  const budgeted = budgetForYear[cat] || 0;
                  const actual = actualByCategory[cat] || 0;
                  if (budgeted === 0 && actual === 0) return null;
                  const pct = budgeted > 0 ? Math.min(100, (actual / budgeted) * 100) : 100;
                  const over = budgeted > 0 && actual > budgeted;
                  return (
                    <div key={cat}>
                      <div className="flex items-center justify-between mb-1">
                        <span className="text-[13px] font-medium">{cat}</span>
                        <span className={`text-[12px] ${over ? "text-red-500" : textMuted}`}>
                          {formatMoney(actual)} {budgeted > 0 ? `/ ${formatMoney(budgeted)}` : "(not budgeted this year)"}
                        </span>
                      </div>
                      <div className={`w-full h-2 rounded-full ${trackBg} overflow-hidden`}>
                        <div
                          className={`h-full rounded-full ${over ? "bg-red-500" : "bg-blue-500"}`}
                          style={{ width: `${pct}%` }}
                        />
                      </div>
                    </div>
                  );
                })}
                {Object.keys(actualByCategory).length === 0 && (
                  <div className={`text-[13px] ${textMuted} text-center py-4`}>
                    Log a few expenses on the Home tab to see them compared against this year's budget.
                  </div>
                )}
              </div>
            </div>

            {/* Full cost head breakdown */}
            <div className={`rounded-2xl border ${border} ${cardBg} overflow-hidden`}>
              <div className="px-6 py-4">
                <span className="font-semibold text-[15px]">Full Cost Head Breakdown — {scenario}</span>
                <div className={`text-[12px] ${textMuted} mt-0.5`}>From your uploaded College_Budget.xlsx — Budget sheet, rescaled to the selected scenario</div>
              </div>
              <div className="overflow-x-auto">
              <table className="w-full text-[13px]">
                <thead>
                  <tr className={`border-t border-b ${border} ${textMuted}`}>
                    <th className="text-left font-medium px-6 py-2.5">Cost Head</th>
                    <th className="text-left font-medium py-2.5">Year {budgetYear}</th>
                    <th className="text-left font-medium py-2.5">5-Year Total</th>
                    <th className="text-left font-medium py-2.5 pr-6">Notes</th>
                  </tr>
                </thead>
                <tbody>
                  {scaledRows.map(row => (
                    <tr key={row.head} className={`border-b ${border} ${rowHover}`}>
                      <td className="px-6 py-2.5">
                        {row.head}
                        {row.variable && <span className={`ml-2 text-[10px] px-1.5 py-0.5 rounded ${dark ? "bg-blue-500/10 text-blue-400" : "bg-blue-50 text-blue-600"}`}>variable</span>}
                      </td>
                      <td className="py-2.5">{formatMoney(row.scaledYears[budgetYear - 1])}</td>
                      <td className={`py-2.5 ${textMuted}`}>{formatMoney(row.scaledYears.reduce((a, b) => a + b, 0))}</td>
                      <td className={`py-2.5 pr-6 ${textMuted}`}>{row.notes}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              </div>
            </div>
          </>
        )}

        {activeView === "analytics" && (
          <>
            {transactions.length === 0 ? (
              <div className={`rounded-2xl border ${border} ${cardBg} p-10 text-center`}>
                <BarChart2 size={28} className={`mx-auto mb-3 ${textMuted}`} />
                <div className="font-semibold text-[15px] mb-1">No transactions yet</div>
                <div className={`text-[13px] ${textMuted} mb-3`}>Log a few entries on the Home tab and trends will show up here.</div>
                <button onClick={loadSampleData} className="text-blue-600 text-[13px] font-medium hover:text-blue-700">
                  Or load sample data to see it in action
                </button>
              </div>
            ) : (
              <>
                {/* Balance trend */}
                <div className={`rounded-2xl border ${border} ${cardBg} p-5 mb-5`}>
                  <div className="font-semibold text-[15px] mb-1">Balance Trend</div>
                  <div className={`text-[12px] ${textMuted} mb-4`}>Running balance after every logged transaction, in order</div>
                  <div style={{ width: "100%", height: 260 }}>
                    <ResponsiveContainer>
                      <LineChart data={balanceTrendData} margin={{ left: 0, right: 12, top: 8, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#262626" : "#e5e5e5"} />
                        <XAxis dataKey="label" tick={{ fontSize: 11, fill: dark ? "#a3a3a3" : "#737373" }} />
                        <YAxis tick={{ fontSize: 11, fill: dark ? "#a3a3a3" : "#737373" }} tickFormatter={v => `${currency.symbol}${(v / 1000).toFixed(0)}k`} />
                        <Tooltip
                          formatter={v => formatMoney(v)}
                          contentStyle={{ background: dark ? "#171717" : "#fff", border: `1px solid ${dark ? "#262626" : "#e5e5e5"}`, borderRadius: 8, fontSize: 12 }}
                        />
                        <Line type="monotone" dataKey="balance" stroke="#3b82f6" strokeWidth={2.5} dot={{ r: 3 }} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-5 mb-5">
                  {/* Spending by category */}
                  <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                    <div className="font-semibold text-[15px] mb-1">Spending by Category</div>
                    <div className={`text-[12px] ${textMuted} mb-4`}>{periodLabel} — per header</div>
                    {periodSpendingChartData.length === 0 ? (
                      <div className={`text-[13px] ${textMuted} py-10 text-center`}>No expenses logged {periodMode === "lifetime" ? "yet" : "in " + periodLabel.toLowerCase()}.</div>
                    ) : (
                      <div style={{ width: "100%", height: Math.max(180, periodSpendingChartData.length * 32) }}>
                        <ResponsiveContainer>
                          <BarChart data={periodSpendingChartData} layout="vertical" margin={{ left: 8, right: 20, top: 4, bottom: 4 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#262626" : "#e5e5e5"} horizontal={false} />
                            <XAxis type="number" tick={{ fontSize: 11, fill: dark ? "#a3a3a3" : "#737373" }} tickFormatter={v => `${currency.symbol}${(v / 1000).toFixed(0)}k`} />
                            <YAxis type="category" dataKey="name" width={150} tick={{ fontSize: 11, fill: dark ? "#a3a3a3" : "#737373" }} />
                            <Tooltip
                              formatter={v => formatMoney(v)}
                              contentStyle={{ background: dark ? "#171717" : "#fff", border: `1px solid ${dark ? "#262626" : "#e5e5e5"}`, borderRadius: 8, fontSize: 12 }}
                            />
                            <Bar dataKey="value" fill="#ef4444" radius={[0, 4, 4, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    )}
                  </div>

                  {/* Income by source */}
                  <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                    <div className="font-semibold text-[15px] mb-1">Income by Source</div>
                    <div className={`text-[12px] ${textMuted} mb-4`}>{periodLabel} — per header</div>
                    {periodIncomeChartData.length === 0 ? (
                      <div className={`text-[13px] ${textMuted} py-10 text-center`}>No income logged {periodMode === "lifetime" ? "yet" : "in " + periodLabel.toLowerCase()}.</div>
                    ) : (
                      <div style={{ width: "100%", height: Math.max(180, periodIncomeChartData.length * 32) }}>
                        <ResponsiveContainer>
                          <BarChart data={periodIncomeChartData} layout="vertical" margin={{ left: 8, right: 20, top: 4, bottom: 4 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#262626" : "#e5e5e5"} horizontal={false} />
                            <XAxis type="number" tick={{ fontSize: 11, fill: dark ? "#a3a3a3" : "#737373" }} tickFormatter={v => `${currency.symbol}${(v / 1000).toFixed(0)}k`} />
                            <YAxis type="category" dataKey="name" width={150} tick={{ fontSize: 11, fill: dark ? "#a3a3a3" : "#737373" }} />
                            <Tooltip
                              formatter={v => formatMoney(v)}
                              contentStyle={{ background: dark ? "#171717" : "#fff", border: `1px solid ${dark ? "#262626" : "#e5e5e5"}`, borderRadius: 8, fontSize: 12 }}
                            />
                            <Bar dataKey="value" fill="#22c55e" radius={[0, 4, 4, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    )}
                  </div>
                </div>

                {/* Category share pie */}
                <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
                  <div className="font-semibold text-[15px] mb-1">Spending Mix</div>
                  <div className={`text-[12px] ${textMuted} mb-4`}>{periodLabel} — share of spend held by each category</div>
                  {periodSpendingChartData.length === 0 ? (
                    <div className={`text-[13px] ${textMuted} py-10 text-center`}>No expenses logged {periodMode === "lifetime" ? "yet" : "in " + periodLabel.toLowerCase()}.</div>
                  ) : (
                    <div style={{ width: "100%", height: 300 }}>
                      <ResponsiveContainer>
                        <PieChart>
                          <Pie data={periodSpendingChartData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={100} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}>
                            {periodSpendingChartData.map((entry, i) => (
                              <Cell key={entry.name} fill={CHART_COLORS[i % CHART_COLORS.length]} />
                            ))}
                          </Pie>
                          <Tooltip
                            formatter={v => formatMoney(v)}
                            contentStyle={{ background: dark ? "#171717" : "#fff", border: `1px solid ${dark ? "#262626" : "#e5e5e5"}`, borderRadius: 8, fontSize: 12 }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  )}
                </div>
              </>
            )}
          </>
        )}

        {activeView === "settings" && (
          <div className="flex flex-col gap-5">
            <div className={`rounded-lg border ${border} ${cardBg} text-[12px] ${textMuted} px-4 py-2.5 flex items-center gap-2`}>
              <span className={`w-1.5 h-1.5 rounded-full ${
                saveStatus === "saving" ? "bg-amber-400" : saveStatus === "error" ? "bg-red-500" : saveStatus === "saved" ? "bg-green-500" : "bg-neutral-300"
              }`} />
              {saveStatus === "saving" ? "Saving your changes…" : saveStatus === "error" ? "Couldn't save just now — your changes are still on screen, will retry shortly." : saveStatus === "saved" ? "Everything below is saved automatically to this browser." : "Loading your saved wallet…"}
            </div>

            {settingsMessage && (
              <div className={`rounded-lg border ${dark ? "border-blue-900 bg-blue-500/10 text-blue-300" : "border-blue-200 bg-blue-50 text-blue-700"} text-[13px] px-4 py-2.5 flex items-center justify-between`}>
                {settingsMessage}
                <button onClick={() => setSettingsMessage("")} className="hover:opacity-70">×</button>
              </div>
            )}

            {restoreConfirming && (
              <div className={`rounded-lg border ${dark ? "border-amber-900 bg-amber-500/10" : "border-amber-200 bg-amber-50"} px-4 py-3 flex items-center justify-between flex-wrap gap-2`}>
                <span className="text-[13px] font-medium text-amber-700">
                  {restoreSource === "cloud"
                    ? "Newer data was found in your cloud sync gist. Load it? This will replace everything currently in the app."
                    : "This backup will replace all current transactions and settings. Continue?"}
                </span>
                <div className="flex items-center gap-2">
                  <button onClick={confirmRestore} className="px-3 py-1.5 rounded-lg text-[13px] font-medium bg-amber-600 text-white hover:bg-amber-700">
                    {restoreSource === "cloud" ? "Load Cloud Data" : "Restore & Overwrite"}
                  </button>
                  <button onClick={cancelRestore} className={`px-3 py-1.5 rounded-lg text-[13px] font-medium border ${border} ${textMuted}`}>
                    Cancel
                  </button>
                </div>
              </div>
            )}

            {/* Account */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="flex items-center gap-3 mb-4">
                <div className={`w-9 h-9 rounded-full flex items-center justify-center ${dark ? "bg-blue-500/15" : "bg-blue-50"}`}>
                  <User size={16} className="text-blue-600" />
                </div>
                <div>
                  <div className="font-semibold text-[14px]">Account</div>
                  <div className={`text-[12px] ${textMuted}`}>Local profile only — no cloud sync yet</div>
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Display name</label>
                  <input
                    type="text"
                    value={accountName}
                    onChange={e => setAccountName(e.target.value)}
                    placeholder="e.g. Aarav"
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] outline-none`}
                  />
                </div>
                <div>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Email (optional)</label>
                  <input
                    type="email"
                    value={accountEmail}
                    onChange={e => setAccountEmail(e.target.value)}
                    placeholder="for future cloud sync"
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] outline-none`}
                  />
                </div>
              </div>
            </div>

            {/* Appearance & Currency */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="font-semibold text-[14px] mb-4">Appearance & Currency</div>
              <div className="flex items-center justify-between mb-4">
                <div>
                  <div className="text-[13px] font-medium">Dark Mode</div>
                  <div className={`text-[12px] ${textMuted}`}>Switch the whole app to a dark theme</div>
                </div>
                <button
                  onClick={() => setDark(d => !d)}
                  className={`w-11 rounded-full relative transition-colors ${dark ? "bg-blue-500" : "bg-neutral-300"}`}
                  style={{ height: 22 }}
                >
                  <span className={`absolute top-0.5 w-4.5 h-4.5 rounded-full bg-white transition-all ${dark ? "left-6" : "left-0.5"}`} style={{ width: 18, height: 18 }} />
                </button>
              </div>
              <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Currency & number format</label>
              <div className="flex items-center gap-2 flex-wrap">
                {CURRENCY_PRESETS.map(c => (
                  <button
                    key={c.label}
                    onClick={() => setCurrency(c)}
                    className={`px-3.5 py-1.5 rounded-lg text-[13px] font-medium border transition-colors ${
                      currency.label === c.label
                        ? "bg-blue-600 border-blue-600 text-white"
                        : `${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`
                    }`}
                  >
                    {c.label}
                  </button>
                ))}
              </div>
            </div>

            {/* Opening balance */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="font-semibold text-[14px] mb-1">Opening Balance</div>
              <div className={`text-[12px] ${textMuted} mb-4`}>Added on top of your logged transactions to calculate your current balance — useful if you're starting mid-way through the year</div>
              <div className="flex items-center gap-3">
                <div className={`flex items-center rounded-lg border ${inputBg} px-3 flex-1 max-w-xs`}>
                  <span className={textMuted}>{currency.symbol}</span>
                  <input
                    type="number"
                    value={openingBalanceInput}
                    onChange={e => setOpeningBalanceInput(e.target.value)}
                    className="w-full bg-transparent outline-none px-2 py-2 text-[14px]"
                  />
                </div>
                <button
                  onClick={applyOpeningBalance}
                  className="px-4 py-2 rounded-lg text-[13px] font-medium bg-neutral-900 text-white hover:bg-neutral-800"
                >
                  Save
                </button>
              </div>
            </div>

            {/* Manage headers */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="font-semibold text-[14px] mb-1">Manage Categories</div>
              <div className={`text-[12px] ${textMuted} mb-4`}>Add or remove the headers that show up in Add Money / Spend Money</div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                <div>
                  <div className="text-[13px] font-medium mb-2">Income sources</div>
                  <div className="flex flex-wrap gap-1.5 mb-3">
                    {sources.map(s => (
                      <span key={s} className={`inline-flex items-center gap-1.5 text-[12px] font-medium px-2.5 py-1 rounded-full ${dark ? "bg-green-500/10 text-green-400" : "bg-green-50 text-green-700"}`}>
                        {s}
                        <button onClick={() => removeSourceHeader(s)} className="hover:opacity-70">×</button>
                      </span>
                    ))}
                    {sources.length === 0 && <span className={`text-[12px] ${textMuted}`}>No sources left.</span>}
                  </div>
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      value={newSourceInput}
                      onChange={e => setNewSourceInput(e.target.value)}
                      onKeyDown={e => e.key === "Enter" && addSourceHeader()}
                      placeholder="New source"
                      className={`flex-1 rounded-lg border ${inputBg} px-3 py-1.5 text-[13px] outline-none`}
                    />
                    <button onClick={addSourceHeader} className="px-3 py-1.5 rounded-lg text-[13px] font-medium bg-green-600 text-white hover:bg-green-700">
                      Add
                    </button>
                  </div>
                </div>
                <div>
                  <div className="text-[13px] font-medium mb-2">Expense categories</div>
                  <div className="flex flex-wrap gap-1.5 mb-3">
                    {categories.map(c => (
                      <span key={c} className={`inline-flex items-center gap-1.5 text-[12px] font-medium px-2.5 py-1 rounded-full ${dark ? "bg-red-500/10 text-red-400" : "bg-red-50 text-red-700"}`}>
                        {c}
                        <button onClick={() => removeCategoryHeader(c)} className="hover:opacity-70">×</button>
                      </span>
                    ))}
                    {categories.length === 0 && <span className={`text-[12px] ${textMuted}`}>No categories left.</span>}
                  </div>
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      value={newCategoryInput}
                      onChange={e => setNewCategoryInput(e.target.value)}
                      onKeyDown={e => e.key === "Enter" && addCategoryHeader()}
                      placeholder="New category"
                      className={`flex-1 rounded-lg border ${inputBg} px-3 py-1.5 text-[13px] outline-none`}
                    />
                    <button onClick={addCategoryHeader} className="px-3 py-1.5 rounded-lg text-[13px] font-medium bg-red-500 text-white hover:bg-red-600">
                      Add
                    </button>
                  </div>
                </div>
              </div>
              <div className={`text-[11px] ${textMuted} mt-3`}>Removing a category some Budget Plan rows map to just hides that comparison — it won't affect your logged transactions.</div>
            </div>

            {/* Export */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="font-semibold text-[14px] mb-1">Export Data</div>
              <div className={`text-[12px] ${textMuted} mb-4`}>Download a copy of your wallet</div>
              <div className="flex items-center gap-3 flex-wrap">
                <button
                  onClick={exportPDFSummary}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                >
                  <FileText size={14} /> Export PDF Summary
                </button>
                <button
                  onClick={exportCSV}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                >
                  <FileSpreadsheet size={14} /> Export CSV
                </button>
                <button
                  onClick={exportExcel}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                >
                  <FileSpreadsheet size={14} /> Export Excel (Ledger + Analysis)
                </button>
                <button
                  onClick={exportJSONBackup}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                >
                  <Download size={14} /> Download JSON Backup
                </button>
              </div>
              <div className={`text-[11px] ${textMuted} mt-3`}>
                The Excel file includes a Ledger sheet plus separate pivot-style analysis sheets (spend by category, income by source, payment method, monthly summary, budget vs actual). These are precomputed summary tables shaped like pivot table output — Excel's own interactive, draggable PivotTable objects aren't something this export can generate, so treat them as a finished report rather than something you rebuild the layout of inside Excel.
              </div>
            </div>

            {/* Cloud Sync */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="flex items-center gap-3 mb-1">
                <div className={`w-9 h-9 rounded-full flex items-center justify-center ${dark ? "bg-blue-500/15" : "bg-blue-50"}`}>
                  <Cloud size={16} className="text-blue-600" />
                </div>
                <div>
                  <div className="font-semibold text-[14px]">Cloud Sync</div>
                  <div className={`text-[12px] ${textMuted}`}>Keep your wallet in sync across devices using a private GitHub Gist</div>
                </div>
              </div>

              <div className={`flex items-center gap-2 text-[12px] ${textMuted} my-3`}>
                <span className={`w-1.5 h-1.5 rounded-full ${
                  cloudSyncStatus === "syncing" || cloudSyncStatus === "connecting" ? "bg-amber-400"
                    : cloudSyncStatus === "error" ? "bg-red-500"
                    : cloudSyncStatus === "synced" ? "bg-green-500" : "bg-neutral-300"
                }`} />
                {cloudSyncStatus === "disabled" ? "Not connected — your data stays local to this browser."
                  : cloudSyncStatus === "connecting" ? "Connecting…"
                  : cloudSyncStatus === "syncing" ? "Syncing…"
                  : cloudSyncStatus === "error" ? "Couldn't reach GitHub — will retry on the next change."
                  : "Synced."}
              </div>

              {!cloudConnected ? (
                <>
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>GitHub Personal Access Token</label>
                  <input
                    type="password"
                    value={githubToken}
                    onChange={e => setGithubToken(e.target.value)}
                    placeholder="ghp_..."
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] mb-3 outline-none`}
                  />
                  <label className={`text-[12px] font-medium ${textMuted} mb-1 block`}>Gist ID (leave blank to create a new one)</label>
                  <input
                    type="text"
                    value={gistId}
                    onChange={e => setGistId(e.target.value)}
                    placeholder="only needed when linking a second device"
                    className={`w-full rounded-lg border ${inputBg} px-3 py-2 text-[14px] mb-4 outline-none`}
                  />
                  <button
                    onClick={connectCloudSync}
                    className="px-4 py-2 rounded-lg text-[13px] font-medium bg-blue-600 text-white hover:bg-blue-700"
                  >
                    Connect & Sync
                  </button>
                </>
              ) : (
                <>
                  <div className="flex items-center gap-2 mb-4 flex-wrap">
                    <span className={`text-[12px] ${textMuted}`}>Gist ID:</span>
                    <code className={`text-[12px] px-2 py-1 rounded ${dark ? "bg-neutral-800" : "bg-neutral-100"}`}>{gistId}</code>
                    <button
                      onClick={() => { navigator.clipboard?.writeText(gistId); setSettingsMessage("Gist ID copied — paste it into your other device's Cloud Sync setup."); }}
                      className={`${textMuted} hover:text-blue-500`}
                      title="Copy Gist ID"
                    >
                      <Copy size={13} />
                    </button>
                  </div>
                  <div className="flex items-center gap-3 flex-wrap">
                    <button
                      onClick={manualSyncNow}
                      className={`px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                    >
                      Sync Now
                    </button>
                    <button
                      onClick={disconnectCloudSync}
                      className={`px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                    >
                      Disconnect
                    </button>
                  </div>
                </>
              )}

              <details className="mt-4">
                <summary className={`text-[12px] font-medium cursor-pointer ${textMuted}`}>How do I set this up?</summary>
                <div className={`text-[12px] ${textMuted} mt-2 space-y-1.5`}>
                  <p>1. On GitHub, go to Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token, and check only the <span className="font-medium">gist</span> scope.</p>
                  <p>2. Paste that token above and click "Connect & Sync" with the Gist ID left blank — this creates a new private gist and shows you its ID.</p>
                  <p>3. On any other device, paste the <span className="font-medium">same token</span> and that <span className="font-medium">Gist ID</span>, then click Connect — it'll offer to load the synced data.</p>
                  <p>Keep the token private — anyone who has it can read or write that gist. Disconnecting here doesn't delete the gist itself, so reconnecting later with the same ID picks up right where you left off.</p>
                </div>
              </details>
            </div>

            {/* Import */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="font-semibold text-[14px] mb-1">Import Data</div>
              <div className={`text-[12px] ${textMuted} mb-4`}>Bring transactions in from a CSV, or restore a full CampCash JSON backup</div>
              <div className="flex items-center gap-3 flex-wrap mb-3">
                <button
                  onClick={triggerCSVImport}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                >
                  <Upload size={14} /> Import Transactions (CSV)
                </button>
                <button
                  onClick={triggerJSONRestore}
                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
                >
                  <Upload size={14} /> Restore Full Backup (JSON)
                </button>
                <input ref={csvFileInputRef} type="file" accept=".csv,text/csv" className="hidden" onChange={handleCSVFileSelected} />
                <input ref={jsonFileInputRef} type="file" accept=".json,application/json" className="hidden" onChange={handleJSONFileSelected} />
              </div>
              <div className={`text-[11px] ${textMuted}`}>
                CSV columns recognized: <span className="font-medium">Date, Type (Income/Expense), Category, Amount, Notes, Payment Method</span> — column order doesn't matter, extra columns are ignored. CSV import adds to your existing ledger; it never removes anything.
              </div>
            </div>

            {/* Sample Data */}
            <div className={`rounded-2xl border ${border} ${cardBg} p-5`}>
              <div className="flex items-center gap-3 mb-1">
                <Sparkles size={16} className="text-blue-500" />
                <div className="font-semibold text-[14px]">Sample Data</div>
              </div>
              <div className={`text-[12px] ${textMuted} mb-4`}>Loads about 3 months of realistic demo transactions — useful for trying out Analytics, the Budget Plan comparison, and month-end rollover without entering anything by hand. Adds to your existing ledger; nothing is overwritten.</div>
              <button
                onClick={loadSampleData}
                className={`px-4 py-2 rounded-lg text-[13px] font-medium border ${border} ${textMuted} hover:${dark ? "bg-neutral-800" : "bg-neutral-50"}`}
              >
                Load Sample Data
              </button>
            </div>

            {/* Reset */}
            <div className={`rounded-2xl border ${dark ? "border-red-900 bg-red-500/5" : "border-red-100 bg-red-50/50"} p-5`}>
              <div className="flex items-center gap-3 mb-1">
                <AlertTriangle size={16} className="text-red-500" />
                <div className="font-semibold text-[14px] text-red-700">Reset / Clear All Data</div>
              </div>
              <div className={`text-[12px] ${textMuted} mb-4`}>Permanently clears every logged transaction, including the saved copy in this browser. Account name, currency, categories, and opening balance are kept.</div>
              {!resetConfirming ? (
                <button
                  onClick={() => setResetConfirming(true)}
                  className="px-4 py-2 rounded-lg text-[13px] font-medium bg-red-500 text-white hover:bg-red-600"
                >
                  Reset All Data
                </button>
              ) : (
                <div className="flex items-center gap-3">
                  <span className="text-[13px] font-medium text-red-700">Are you sure? This can't be undone.</span>
                  <button onClick={handleResetAll} className="px-3 py-1.5 rounded-lg text-[13px] font-medium bg-red-600 text-white hover:bg-red-700">
                    Confirm Reset
                  </button>
                  <button onClick={() => setResetConfirming(false)} className={`px-3 py-1.5 rounded-lg text-[13px] font-medium border ${border} ${textMuted}`}>
                    Cancel
                  </button>
                </div>
              )}
            </div>
          </div>
        )}
      </main>

      {lastDeleted && (
        <div className={`fixed bottom-5 left-1/2 -translate-x-1/2 z-30 rounded-lg border ${border} ${cardBg} shadow-lg px-4 py-3 flex items-center gap-4`}>
          <span className="text-[13px]">Transaction deleted.</span>
          <button onClick={undoDelete} className="text-[13px] font-semibold text-blue-600 hover:text-blue-700">
            Undo
          </button>
        </div>
      )}
    </div>
  );
}
