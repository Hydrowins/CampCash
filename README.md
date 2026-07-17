# CampCash

A personal budgeting and wallet tracker — built for everyday spending, fixed deposits, and a real multi-year budget plan, all in one dashboard.

> **A note on how this was made:** The UI/UX design was created by hand in **Figma**. All of the actual coding — every component, calculation, and feature below — was written by **Claude AI**, iterated on conversationally rather than typed by hand line by line.

**🔗 Live demo:** [camp-cash-sigma.vercel.app](https://camp-cash-sigma.vercel.app/)

---

## What it does

### 💰 Wallet & Transactions
- Add money / spend money with a live running balance
- Custom, editable categories and income sources — not locked to a fixed list
- Per-expense payment method tagging (Cash / Card / UPI / Other)
- Quick-add chips for common amounts
- Full transaction ledger with search, category filtering, and sortable columns (date, amount)
- Undo on delete, so nothing's ever gone by accident

### 📅 Time-Aware by Design
- Automatic **month-end rollover** — your balance carries forward into the next month with a visible, auditable ledger entry, just like a bank statement
- A time-period selector (specific month, or lifetime) that reshapes the dashboard, analytics, and ledger to match

### 🏦 Fixed Deposits
- Track FDs with principal, rate, tenure, and compounding frequency (monthly / quarterly / annually / simple)
- Live maturity date and maturity amount calculations, with a progress bar toward maturity
- Optional automatic wallet debit when you open an FD, and automatic credit when you mark it matured or withdraw early — both fully opt-out if you're just tracking

### 📊 A Real Budget Plan
- A full 5-year budget, broken down by cost head and year, imported directly from a spreadsheet
- Conservative / Comfortable / High scenario multiplier, applied only to the actual variable, personal costs — fixed costs stay fixed
- "Budgeted vs Logged" comparison, showing how real spending stacks up against the plan

### 📈 Analytics
- Balance trend over time
- Spending by category and income by source, both period-aware
- Spending mix breakdown
- All built with live charts, not static images

### 🔔 Notifications
- Over-budget alerts per category
- Low-balance warnings
- Reminders when a Fixed Deposit has matured but hasn't been closed out yet

### 📤 Export & Import
- Export your full ledger as **CSV**
- Export a multi-sheet **Excel workbook** — ledger plus pivot-style analysis sheets (spend by category, income by source, payment method, monthly summary, budget vs actual)
- Export a clean **PDF summary**, print-ready
- Import transactions from **CSV**, or restore a complete **JSON backup**
- One-click **sample data** to explore the app without entering anything by hand

### ☁️ Cloud Sync
- Optional sync across devices using a private **GitHub Gist** — no paid backend, no accounts, no billing setup
- Local-first: works fully offline with browser storage if sync is never turned on

### 🎨 Polish
- Full dark mode
- Mobile-responsive layout, including a dedicated bottom navigation bar
- Persistent local storage, so your data survives a refresh even without cloud sync

---

## Tech Stack

- **React** — component structure and state management
- **Tailwind CSS** — styling
- **Recharts** — charts and data visualization
- **SheetJS (xlsx)** — Excel export with multiple sheets
- **Lucide** — icon set

---

## Try It Yourself

You've got two paths to get this running — pick whichever fits.

### Option A: StackBlitz (no installs, runs in your browser)

The fastest way to poke around or fork your own copy.

1. Go to [stackblitz.com](https://stackblitz.com) and start a new **React** project (search "react" in the template picker, or go straight to `stackblitz.com/fork/react`)
2. In the file panel, create a new file `src/campcash-app.tsx` and paste in the app's code
3. Replace `src/App.tsx` with:
   ```tsx
   import CampCash from './campcash-app';
   export default function App() {
     return <CampCash />;
   }
   ```
4. Open the terminal at the bottom and install the libraries the app needs:
   ```bash
   npm install lucide-react recharts xlsx
   ```
5. Install Tailwind — **use v3, not the current default v4**, since v4's setup is a different, incompatible flow:
   ```bash
   npm install -D tailwindcss@3 postcss autoprefixer
   npx tailwindcss init -p
   ```
6. In `tailwind.config.js`, set:
   ```js
   content: ["./index.html", "./src/**/*.{js,jsx,ts,tsx}"],
   ```
7. In `src/index.css`, replace everything with:
   ```css
   @tailwind base;
   @tailwind components;
   @tailwind utilities;
   ```
8. **Restart the dev server** (`Ctrl+C` in the terminal, then `npm run dev` again) — Tailwind's config often doesn't take effect on a hot-reload alone, only on a fresh boot.

If a paste into the terminal seems to do nothing when you hit Enter, click directly into the terminal first (`Ctrl+C` to clear it, then type the command) — StackBlitz's terminal occasionally needs focus reasserted before it accepts input.

### Option B: Locally, with VS Code

1. Install [Node.js](https://nodejs.org) (LTS version)
2. Scaffold a project:
   ```bash
   npm create vite@latest campcash -- --template react
   cd campcash
   npm install
   ```
3. Install dependencies:
   ```bash
   npm install lucide-react recharts xlsx
   npm install -D tailwindcss@3 postcss autoprefixer
   npx tailwindcss init -p
   ```
4. Configure `tailwind.config.js` and `src/index.css` exactly as in steps 6–7 above
5. Drop `campcash-app.jsx` into `src/`, and point `App.jsx` at it the same way as step 3 above
6. Run it:
   ```bash
   npm run dev
   ```

### Deploying to Vercel (a real, always-on URL)

Once your code is pushed to a GitHub repo (StackBlitz has a built-in "Create a repository" option in the sidebar, or push manually from a local clone):

1. Go to [vercel.com](https://vercel.com) → sign in with GitHub
2. **Add New... → Project**, then import your repo
3. Vercel auto-detects Vite — leave the default build settings
4. Before deploying, open `package.json` and make sure the build script is:
   ```json
   "build": "vite build",
   ```
   not `"tsc -b && vite build"` — the stricter TypeScript type-check isn't needed here and will fail the build over things that don't actually affect the app (missing type annotations on plain JSX, mostly). You can edit this straight on GitHub if it's easier than through StackBlitz.
5. Click **Deploy**

A minute later you'll have a live `.vercel.app` URL — no server to keep running, no tab that needs to stay open, works from any device. Any future commit to the repo redeploys it automatically.

---

## Status

CampCash is under active iteration — categories, budget structure, and sync are all designed to flex as real usage uncovers what's actually needed next.
