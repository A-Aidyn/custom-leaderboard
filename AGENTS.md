# AGENTS.md

## Project Overview
Google Apps Script (GAS) project that calculates Elo-style player ratings for Valorant custom matches. Runs inside Google Sheets via `SpreadsheetApp`. No local build/test/lint toolchain — deploy via the Apps Script editor.

## Commands
- **Run**: Execute `calculateAllRatings` from the Apps Script editor or the Sheets menu ("Refresh Leaderboard")
- **No local tests or linter** — validate in the Apps Script environment; check `Logger.log` output.

## Architecture
- **code.js** — Single-file project:
  - `validateMatch` — validates team labels on match player rows.
  - `calculateAllRatings` — reads "Raw Data" sheet, computes modified Elo ratings (ACS/KDA performance index), writes ranked results to "Leaderboard" with tier-colored conditional formatting.
  - `onOpen` — adds a custom menu to the spreadsheet UI.
- **Data**: Google Sheet with "Raw Data" (MatchID, Player, Team, RoundsWon/Lost, ACS, K/D/A) and "Leaderboard" (output).

## Code Style
- Plain ES5-compatible JS (GAS runtime). No modules, imports, or TypeScript.
- `Logger.log` for debug; `SpreadsheetApp.getUi().alert` for user messages.
- Constants (K-factor, base rating 1500) are inline. Use descriptive names and inline comments for math.
- 2-space indentation. Prefer semicolons consistently.
