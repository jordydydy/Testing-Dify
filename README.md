# Procurement Test Runner

Local web UI that runs `Test_Document_Procurement.xlsx` against your Dify app and tells you, per test, whether the
agent cited the expected document. Built for repeated runs: change a prompt or workflow, click Run, compare.

## Start
Double-click `run_tester.bat` (or `python procurement_tester.py`). The browser opens at http://127.0.0.1:8765.
Needs Python 3 with `requests` and `openpyxl` (the .bat installs them if missing). The server listens on localhost only.

## Auto-start at login
Already installed: a shortcut in your Windows Startup folder (`ProcurementTestRunner.lnk`) starts the runner hidden every
time you log in. No console window and no browser popup: just open http://127.0.0.1:8765 (bookmark it).
- Re-install or start it now: `autostart_install.bat`
- Remove it and stop the runner: `autostart_remove.bat`
- Log file: `test_runner.log`. A second launch exits by itself because the port is already in use.
- It starts when you log in to Windows (not before the login screen) and keeps running until you sign out.
- After editing `procurement_tester.py`, restart it: run `autostart_remove.bat`, then `autostart_install.bat`. Changes to `ui.html` only need a page refresh.

## Use
1. **Connection**: Mode, Base URL (with `/v1`), API key, then **Test connection**.
2. **Test file**: the default path is already filled. Click **Load**, or upload another `.xlsx` in the same format.
   **Download blank template** gives the static template (`Test_Template.xlsx`, also saved in this folder).
3. **Run**: *Run selected*, *Run all auto-checkable*, or *Re-run failed / error*. Click any row for the full detail.
4. **Compare**: pick a previous run in "compare with" to see `PASS -> FAIL` regressions per test.
5. **Export Excel**: writes a COPY of your workbook into `results/` with Actual Result, Evidence, Execution Time,
   Status and Comments filled in. Your original file is never modified. CSV and JSON are also available.

## TLS / "CERTIFICATE_VERIFY_FAILED"
Corporate networks re-sign HTTPS with an internal CA, and Python's bundled certificate list does not know it. The **TLS certificates** setting:
- **Windows trusted certificates (default, recommended)**: full verification against the Windows Root store, accepting an intermediate CA that IT
  installed as a trust anchor (what browsers do). This fixed `genaistudio.telkomsel.co.id`: IT installed "Telekomunikasi Selular Pt Forward Trust CA" in
  the Windows Trusted Root store, but the server sends it signed by "Pt Root CA", which is not installed, so strict Python/OpenSSL rejected the path.
- **Custom CA file**: give a .pem/.crt path (for example a root exported by IT). Still fully verified.
- **Do not verify**: turns certificate checking off. Last resort only; the API key is then sent over a connection that is not authenticated.

Note: the "Forward Trust CA" certificates in the Windows store expire (2 of 4 expired 2026-06-28, the other 2 on 2028-06-26). If the error returns after a
rotation, ask IT to install the new certificate.

## Modes
| Mode | Calls | Use it to |
|---|---|---|
| Dify Agent / Chat - streaming | `POST /chat-messages` (SSE) | Test the real agent. Required for Dify **Agent** apps. |
| Dify Chat - blocking | `POST /chat-messages` | Chatflow apps that allow blocking. |
| Dify Workflow (tool only) | `POST /workflows/run` | Test retrieval only: did the tool return the expected document, at which level? |
| Custom HTTP endpoint | your URL, body template, answer JSON path | Portal endpoints. |

Run the same questions in agent mode and tool-only mode. If a test FAILS in agent mode but the tool returned the
expected document, the agent's judgement is wrong. If the tool did not return it, retrieval is wrong. The runner
prints this diagnosis in the Reason column ("agent judgement" vs "retrieval").

## The static template
Same 16 columns as the project workbook, plus two optional columns at the end:
- **Expected Document**: exact file name the answer must cite. Several acceptable documents: separate with `|`.
- **Test Type**: `document`, `fallback`, `security` or `manual`.

If they are empty, the runner guesses from the title / prerequisite / expected result (this is how the current workbook works).
A second sheet, "How to fill", documents every column. Existing workbooks keep working unchanged.

## What counts as PASS
| Test kind | PASS when |
|---|---|
| document (76 rows) | the first `Document source:` in the answer is the expected document. Several accepted documents are supported (for example a regulation and its amendment). Names are matched ignoring case, extension, and punctuation. |
| fallback (2) | no `Document source:` line, and the apology or the escalation emails are present |
| security (1) | a refusal and no `Document source:` line |
| manual (1) | never auto-passed: shown as MANUAL |
| blank rows (3) | skipped |

Also recorded for every test: total time (and first-token time), the query the agent wrote, how many tool calls it made
(more than 1 is flagged, the rule is exactly one), which documents the tool returned and at which level, and a content
score (share of the expected excerpt's distinctive words found in the answer; informational, or set a minimum).

## Settings worth knowing
- **Parallel workers** (default 3): faster. Timings then include server queueing. Use 1 for a clean latency measurement.
- **Time limit** (default 50 s): tests over it are flagged SLOW; tick "fail tests slower than the limit" to make it a failure.
- **New conversation per test**: always. Each test starts with an empty conversation.
- **Clean test steps**: removes "Halo min", "Chatbot Answer" and quotes from the questions. Rows with
  `input channel_id: x` always have that first line removed from the question.
- **Remember settings**: stores the settings, including the API key, in this browser's local storage only.

## Files
- `procurement_tester.py` server, Dify calls, evaluation, export
- `ui.html` the page (edit freely; it is re-read on every page load)
- `test_runs/` every run as JSON (feeds "compare with")
- `results/` exported Excel / CSV / JSON
- `uploads/` workbooks uploaded through the page

## Limits
- It judges the cited document, not the quality of the prose. Read the content score and the full answer for that.
- The workbook covers 14 of the 42 documents in the knowledge base.
