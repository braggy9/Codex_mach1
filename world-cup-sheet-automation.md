# World Cup tipping sheet automation

This guide describes a low-maintenance way to keep a Google Sheet updated with World Cup fixture scores so you can compare live results against everyone’s tips.

## Recommended architecture

Use **Google Apps Script attached to the spreadsheet** as the automation layer:

1. A `Tips` sheet stores each participant’s predictions.
2. A `Matches` sheet stores one row per fixture and is updated from a football results API.
3. A `Leaderboard` sheet uses formulas or Apps Script to calculate points from `Tips` versus `Matches`.
4. A time-driven Apps Script trigger runs every few minutes during match windows and less often outside matches.

This keeps credentials inside Google, avoids running a server, and writes directly to the sheet.

## Sheet layout

Create these tabs.

### `Config`

| Key | Value |
| --- | --- |
| API_PROVIDER | api-football |
| API_BASE_URL | https://v3.football.api-sports.io |
| API_LEAGUE_ID | 1 |
| API_SEASON | 2026 |
| POLL_MINUTES | 5 |

Store the API key in Apps Script **Script Properties**, not in a visible cell.

### `Matches`

| match_id | utc_kickoff | status | home_team | away_team | home_score | away_score | elapsed | last_updated |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |

`match_id` must be the provider fixture ID. Keep it stable because your formulas and updates depend on it.

### `Tips`

| participant | match_id | home_tip | away_tip |
| --- | --- | --- | --- |

Use data validation on `match_id` from `Matches!A:A` to prevent typos.

### `Leaderboard`

At minimum, calculate:

- exact score points
- correct result points
- total points by participant

Example scoring formula for each tip row:

```gs
=IF(OR(VLOOKUP(B2,Matches!A:G,6,FALSE)="",VLOOKUP(B2,Matches!A:G,7,FALSE)=""),0,
  IF(AND(C2=VLOOKUP(B2,Matches!A:G,6,FALSE),D2=VLOOKUP(B2,Matches!A:G,7,FALSE)),3,
    IF(SIGN(C2-D2)=SIGN(VLOOKUP(B2,Matches!A:G,6,FALSE)-VLOOKUP(B2,Matches!A:G,7,FALSE)),1,0)))
```

## Apps Script setup

1. Open the Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Add the code below.
4. In **Project Settings → Script Properties**, add `FOOTBALL_API_KEY`.
5. Run `syncWorldCupFixtures` manually once and approve permissions.
6. Run `installTriggers` once.

Google Apps Script’s `UrlFetchApp` is the service used to call external HTTP APIs, and `PropertiesService` is used to keep secrets out of the sheet.

## Apps Script example

This example uses API-Football / API-Sports for the 2026 FIFA World Cup. If you use another provider, keep the same `Matches` tab schema and only replace `fetchFixtures_` and `normaliseFixture_`.

```javascript
const SHEET_MATCHES = 'Matches';
const SHEET_CONFIG = 'Config';

function installTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'syncWorldCupFixtures')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('syncWorldCupFixtures')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function syncWorldCupFixtures() {
  const spreadsheet = SpreadsheetApp.getActive();
  const config = readConfig_(spreadsheet.getSheetByName(SHEET_CONFIG));
  const matchesSheet = spreadsheet.getSheetByName(SHEET_MATCHES);
  const fixtures = fetchFixtures_(config);
  const rows = fixtures.map(normaliseFixture_);

  upsertMatches_(matchesSheet, rows);
}

function fetchFixtures_(config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('FOOTBALL_API_KEY');
  if (!apiKey) {
    throw new Error('Missing FOOTBALL_API_KEY in Script Properties.');
  }

  const baseUrl = config.API_BASE_URL || 'https://v3.football.api-sports.io';
  const leagueId = config.API_LEAGUE_ID || '1';
  const season = config.API_SEASON || '2026';
  const url = `${baseUrl}/fixtures?league=${encodeURIComponent(leagueId)}&season=${encodeURIComponent(season)}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'x-apisports-key': apiKey },
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(`Football API request failed: HTTP ${statusCode} ${response.getContentText()}`);
  }

  const payload = JSON.parse(response.getContentText());
  return payload.response || [];
}

function normaliseFixture_(fixture) {
  return [
    fixture.fixture.id,
    fixture.fixture.date,
    fixture.fixture.status.short,
    fixture.teams.home.name,
    fixture.teams.away.name,
    fixture.goals.home,
    fixture.goals.away,
    fixture.fixture.status.elapsed,
    new Date().toISOString(),
  ];
}

function upsertMatches_(sheet, rows) {
  const headers = ['match_id', 'utc_kickoff', 'status', 'home_team', 'away_team', 'home_score', 'away_score', 'elapsed', 'last_updated'];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const existingValues = sheet.getDataRange().getValues();
  const rowByMatchId = new Map();
  existingValues.slice(1).forEach((row, index) => {
    if (row[0]) {
      rowByMatchId.set(String(row[0]), index + 2);
    }
  });

  const newRows = [];
  rows.forEach(row => {
    const existingRowNumber = rowByMatchId.get(String(row[0]));
    if (existingRowNumber) {
      sheet.getRange(existingRowNumber, 1, 1, headers.length).setValues([row]);
    } else {
      newRows.push(row);
    }
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }
}

function readConfig_(sheet) {
  return Object.fromEntries(
    sheet.getDataRange().getValues()
      .slice(1)
      .filter(row => row[0])
      .map(row => [String(row[0]).trim(), row[1]])
  );
}
```

## Operational tips

- **Provider choice:** API-Football documents World Cup 2026 with league `1` and season `2026`. Football-Data.org is another option, but you should confirm its World Cup competition code and plan limits before building formulas around it.
- **Polling frequency:** use five-minute polling for a casual tipping sheet. Only poll every one minute if your API plan allows it.
- **Avoid overwriting manual fields:** keep API-managed columns in `Matches` separate from manual notes or admin override columns.
- **Handle rate limits:** if your provider returns HTTP `429`, increase `POLL_MINUTES` and reduce trigger frequency.
- **Audit updates:** keep `last_updated` visible so you can see whether the sheet is stale.
- **Live scoring:** let formulas calculate points from the latest `Matches` values. Apps Script should focus on reliable data capture.

## Testing checklist

1. Put a test API key in Script Properties.
2. Run `syncWorldCupFixtures` from Apps Script.
3. Confirm `Matches` gets fixture IDs, teams, kickoff times, status, and scores.
4. Change one `Tips` row and confirm `Leaderboard` updates.
5. Temporarily change the trigger to one minute and confirm `last_updated` changes.
6. Restore the trigger to five minutes or your provider’s safe polling interval.

## Optional enhancements

- Add an `Admin` menu with a manual “Sync now” button.
- Add conditional formatting for live matches where `status` is `1H`, `HT`, `2H`, `ET`, or `P`.
- Send Slack, email, or Google Chat notifications when matches finish.
- Cache the last API payload in Script Properties to avoid rewriting unchanged rows.
- Lock protected ranges on `Matches` so participants cannot edit live scores.
