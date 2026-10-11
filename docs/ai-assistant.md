# AI assistant (Beta)

The AI assistant works with the whole workbook, not just the cells on screen. Ask it a question or describe a job, and it reads what it needs: sheets, formulas, charts, highlight rules and dropdowns. When it wants to change something, it proposes the change and shows you exactly what will happen first. Nothing in the workbook changes until you click **Apply**, and every applied change undoes in one click.

> The assistant is in development. It's available in the desktop app on the **Developer branch** (Settings → Updates → Developer branch).

## Getting started

1. Open **Home → Ask AI**, or press **Ctrl+Shift+A**.
2. Connect a model, either:
   - **Claude** — click **Add API key** and paste a key from [console.anthropic.com](https://console.anthropic.com/settings/keys). Usage is billed to your Anthropic account.
   - **A model on this computer** — run [Ollama](https://ollama.com) or [LM Studio](https://lmstudio.ai). Then in **Settings → AI assistant**, set **Use** to *A model on this computer* and click **Find models**. Choose a model that supports tool calling (for example Qwen 3 or Llama 3.1). With a local model, the workbook never leaves your computer.
3. Type a request and press **Enter**.

Things to try:

- "Find duplicate transactions, but don't delete anything."
- "Build a monthly budget using these columns."
- "Explain why my revenue total doesn't match the report."
- "Create a chart showing which products are losing money."
- "Add a formula that categorizes these transactions."

## What you see in the chat

- **Reads.** Every time the assistant looks at the workbook, a line shows what it looked at. For example: *Read Transactions!A4:P200 — 3,200 cells*, *Calculated =SUMIFS(…)*, *Traced Report!D14* or *Searched formulas for "Data!"*.
- **Answers.** Cell addresses in answers, like Transactions!L5, are links. Click one to jump to those cells.
- **Proposals.** A card describes each change the assistant wants to make:
  - what it does, step by step;
  - **every cell that changes**, before → after, including the formula it writes;
  - **other results that change**: formulas elsewhere in the workbook whose values would change as a result, before → after;
  - how many cells only get new formatting.

  The preview is exact. Gridline tries the change on the workbook, measures the result, and puts everything back before you see the card.

## Apply, reject, undo

- **Apply** makes the change as a single step. Formulas recalculate and Gridline jumps to the first changed cell.
- **Reject** leaves the workbook untouched. To tell the assistant what to do instead, type in the box under the buttons and press **Enter**, for example "use a new sheet instead".
- **Undo** on an applied card reverses everything that change did, in one click. This works even after you've made other edits since. Cells you changed yourself afterwards are left alone.
- AI changes are ordinary undo steps too, named like *AI: Add a Category column*, so **Ctrl+Z** and **Ctrl+Y** also work. The card keeps up with them.

## What it can do

The assistant reads with these tools:

| Tool | What it does |
| --- | --- |
| Read cells | Values or formulas of any range, on any sheet |
| Search | Finds text in cells or in formulas, across all sheets |
| Describe a sheet | For each column: the heading, type of data, min/max/sum, blanks and errors. Also the sheet's formulas, charts, highlight rules and dropdowns |
| Trace a cell | A cell's formula, the cells it uses (with their values) and the formulas that use it, across sheets |
| Calculate | Runs any formula (SUMIFS, COUNTIFS, FILTER, UNIQUE, XLOOKUP…) without writing it anywhere. The assistant uses this for totals and checks instead of doing the math itself |

Changes it can propose:

- write values or formulas;
- fill a formula down a range;
- clear cells;
- insert or delete rows and columns;
- add or rename sheets;
- format: bold, colors, number formats, borders, alignment and column widths;
- add charts;
- highlight with conditional formatting (duplicates, greater than, top 10, a formula…);
- sort, add filters, freeze panes;
- add notes and dropdown lists.

It doesn't run macros, open other files or reach the internet.

## Privacy

- With **Claude**, your request and the parts of the workbook the assistant reads are sent to Anthropic. Every read is listed in the chat, so you can see what was sent. Each request also includes a short overview of the workbook: sheet names and sizes, headings, the formulas in use, charts and rules.
- **Settings → AI assistant → Most cells per read** limits how much one read can include (default 2,000 cells).
- With a **local model**, nothing leaves your computer.

### Your API key stays secret

- The key is encrypted with your computer's keychain (Windows Data Protection) and stored in Gridline's settings folder. If that file is copied to another computer or user account, it can't be decrypted there.
- Once saved, the key is never shown again, and the Gridline window itself can't read it back. Only Gridline's background process uses it, to sign requests to Anthropic.
- The key is only sent to `api.anthropic.com`, over HTTPS. If that address ever answered with a redirect, Gridline would stop rather than send the key anywhere else.
- Error messages are checked, and anything that looks like a key is hidden before it's shown.
- The key is never part of what the AI sees, and it isn't included in **Settings → Export**.
- If you paste an API key into the chat by mistake, the message isn't sent.
- If the workbook contains something that looks like an API key (Anthropic, OpenAI, GitHub, AWS, Google, Slack or Stripe), the AI sees `[hidden API key]` instead.
- **Settings → AI assistant → Remove** deletes the key from the computer.

## Settings

**Settings → AI assistant**:

| Setting | What it does |
| --- | --- |
| Use | Claude (your API key) or a model on this computer |
| Claude model | Sonnet 5.5 (recommended), Opus 5.5 (most capable) or Haiku 5.5 (fastest and cheapest) |
| Claude API key | Add, replace or remove the key |
| Server address / Model | Your local server, such as `http://localhost:11434/v1` for Ollama or `http://localhost:1234/v1` for LM Studio, and the model to use |
| Most cells per read | How much one read can include |
