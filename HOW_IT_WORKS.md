# How it works

The repo holds two things. **The detective** is what you download: you point it at a spreadsheet and it explains what the spreadsheet really does. **The exam room** is how we measured the detective: practice workbooks with planted problems, a private answer key and a lock, so its scores are honest. You never need the exam room to use the detective.

## The detective

| Piece | Simple version | Why it's there |
|---|---|---|
| Works on a copy | Never touches your original | You're handing it a real finance file |
| Inventory | Turns out every pocket | Problems hide in hidden sheets, white text, macros and dead links |
| Map | Draws the wiring | You can't explain results without knowing which cells are inputs and which are outputs |
| Probe | Changes one input at a time in real Excel and watches what moves | Reading formulas isn't enough. Excel does surprising things (lookups on unsorted lists, `>` versus `>=` at a boundary), and only real Excel shows what really happens |
| SEEN / INFERRED / UNKNOWN | Facts, guesses and open questions, kept apart | So nobody mistakes a guess for a fact. The unknowns become your meeting agenda |
| Rebuild + Prove | Rewrites the logic in Python and checks it against Excel on hundreds of inputs | The exam the detective sets itself: if the copy matches, it understood the sheet |
| `excelfn` helpers | Excel's odd habits, written in Python | Excel rounds and searches in its own way. Without these, a correct rebuild still wouldn't match |
| `try` | Tests one specific hunch | Turns "I think" into "I checked" |
| Watchdog | Notices when Excel freezes on a pop-up | Otherwise the tool hangs silently |
| One job at a time | Only one stage drives Excel at once | Two stages at once corrupt each other's results |
| Reopen after a stuck error | Starts from a fresh copy if Excel gets stuck | Errors stick inside circular formulas and would give wrong answers |
| LibreOffice fallback | Works without Excel, with caveats | Not everyone has Excel |
| Standard vs exhaustive | A quick check or an audit-grade run | A time-versus-thoroughness choice: standard probes boundaries on every input and samples the rest; exhaustive probes everything and proves on 1,000 cases |

## The exam room

| Piece | Simple version | Why |
|---|---|---|
| 3 sample workbooks + answer key | The practice test. The key stays private | Real scores, and a key that can still be used for blind runs |
| Sealed settings | The answer key is locked away from the session being scored | So the AI can't peek |
| Sandbox | A stronger lock | The AI can read files through code, which permission rules alone don't stop |
| Bridge (`sheetre/bridge.py`) | Lets the locked-in session still drive Excel | The macOS sandbox blocks Excel automation. Needed only because of the lock |
| Read-only tools during the test | The session can't change its own tools | In an early test it tried to |

Someone using the detective on their own workbook needs none of the exam room: no sandbox, no bridge, no answer key.

## What the proof does and doesn't show

Prove compares the rebuild with Excel on every output, case by case: the workbook's saved inputs, 300 random input sets, and a replay of every probe run. A value passes if it matches within 0.000001, or within one part in a billion for large totals. A deliberately broken rebuild, with every margin 0.4 percentage points too high, fails it.

The proof shows the rebuild reproduces the workbook, bugs included. It doesn't show the workbook is right. That's what the ledger is for.
