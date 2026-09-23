# Login works on the original account but not on an imported copy

## Symptom

> I imported this same project into another Gmail account and connected it to the same
> sheet. When I log in on the other account it does not work, but on the original account
> where the project was built it works.

## Root cause: the password salt lived in the Apps Script *project*, not in the spreadsheet

The app does not store your password. It stores `SHA256(password + salt)` in column **H**
of the `Users` sheet, and the salt came from `getSalt()` in `Security.gs`:

```js
// BEFORE
function getSalt() {
  const props = PropertiesService.getScriptProperties();   // <-- per PROJECT
  let salt = props.getProperty('PASSWORD_SALT');
  if (!salt) { salt = Utilities.getUuid(); props.setProperty('PASSWORD_SALT', salt); }
  return salt;
}
```

`Script Properties` belong to the **Apps Script project**, not to the spreadsheet. When you
import/copy the project into another Google account it is a *different* project, so it starts
with **empty script properties**. The first time it runs it silently generates a **new random
salt**, and every hash already stored in the `Users` sheet stops matching:

```
original project : SHA256("admin123" + SALT_A)  ->  stored in the sheet
imported copy    : SHA256("admin123" + SALT_B)  ->  compared with the stored value  ->  FAIL
```

The sheet is shared, but the salt was not. That is why the same username/password works on the
original account and fails on the imported one — for **every** user, not just one.

Two related traps made it worse:

* **Two functions named `verifyPassword`** existed in the same project (`Code.gs` and
  `Security.gs`). Apps Script puts every `.gs` file in one global namespace, so the winner
  depended on file order — and the two implementations understood different password formats.
* **Login crashed instead of failing cleanly** if the account running the script could not
  write to the sheet: clearing the failed-attempt counter threw before the session was created.

## What was changed

| File | Change |
| --- | --- |
| `Security.gs` | The salt is now stored **in the spreadsheet** (hidden `_SystemConfig` sheet), so it travels with the data and is shared by every copy of the project bound to that file. Script Properties are still read/written for backwards compatibility, and the spreadsheet value always wins. |
| `Security.gs` | One canonical `verifyPassword()` that understands all formats the project has ever written: salted SHA-256, legacy `sha256:<hex>`, unsalted SHA-256 and plain text. Also accepts hashes written with a *previous* salt, so nothing breaks during migration. |
| `Code.gs` | Removed the duplicate `verifyPassword()` (kept a comment pointing at `Security.gs`). |
| `User.gs` | Login no longer throws when the sheet is read-only; legacy password formats are silently upgraded to the current salt on the next successful sign-in. |
| `CrossAccountLoginFix.gs` | New: diagnostics + repair tools (`diagnoseLoginEnvironment`, `publishPasswordSaltToSheet`, `syncPasswordSaltFromSheet`, `exportPasswordSalt`, `setPasswordSalt`, `findWorkingSaltFor`, `applySaltForUser`, `tryPasswordSalt`, `setSpreadsheetId`, `showLoginDiagnostics`). |
| `Code.gs` (`onOpen`) | Adds a **🔑 Login Fix** menu to the spreadsheet with the three tools you need. |

Existing passwords keep working — nothing is re-hashed and nobody is forced to reset.

## Fix your two projects (do it in this order)

### 1. Original project (the one where login still works)

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Paste/import the updated `Security.gs`, `Code.gs`, `User.gs` and add the new
   `CrossAccountLoginFix.gs`.
3. In the editor, run **`publishPasswordSaltToSheet`** (select the function → **Run**).
   This writes the salt into the hidden `_SystemConfig` sheet.
   Expected result: `{ success: true, fingerprint: "36:xxxxxx…" }`
4. **Deploy → Manage deployments → Edit → Version: New version.**
   A web app only serves the code that was deployed; if you skip this, the `/exec` URL keeps
   running the old version.

### 2. Imported copy (the other Gmail account)

1. Open the **same spreadsheet** → **Extensions → Apps Script** (this is what makes the
   project container-bound; a standalone project cannot see the sheet).
2. Paste/import the same updated files.
3. Run **`syncPasswordSaltFromSheet`**. Expected result: `success: true`.
4. **Deploy → New deployment → Version: New version** (Web app, Execute as: *Me*,
   Who has access: *Anyone*).
5. Sign in with the same username/password as before. It now works.

### Verify at any time

Run **`diagnoseLoginEnvironment`** (or use the spreadsheet menu
**🔑 Login Fix → Run login diagnostics**) and read the log. It reports:

* whether this project is bound to a spreadsheet and which file it is pointing at;
* whether the account running the script can edit that file;
* where the salt comes from, who published it and when;
* how many users exist and which password format each one uses;
* whether a specific username/password verifies, and with which salt.

Optional: `testUserCredential("username", "password")` returns
`{ passwordMatches: true, matchMode: "salted-current", … }`.

## If it is already broken and you cannot log in anywhere

1. On the **original** project, sign in as Admin and run **`exportPasswordSalt`** in the
   editor. It returns every salt that project knows.
2. On the **imported** copy, run **`tryPasswordSalt("<value from step 1>", "username", "password", true)`**
   for each value. When one matches it is applied automatically and printed as
   `{ matched: true, applied: true }`.
3. If none matches, run **`findWorkingSaltFor("username", "password")`** — it tells you
   whether the credential already verifies (then the problem is sessions/permissions, not the
   salt) or whether the salt really is missing from this copy.

## Other things that break an imported copy (check these too)

* **Deployment version.** The `/exec` URL serves the version that was deployed, not the code
  currently in the editor. Always create a **new version** after pasting code.
* **Execute as / access.** Deploy with **Execute as: Me (the deployer)** and
  **Who has access: Anyone** (this is what `appsscript.json` declares). With "Execute as: the
  user accessing the web app", every visitor must be able to edit the spreadsheet *and* must
  authorise the script themselves.
* **Sharing.** The account the script runs as must be able to **edit** the spreadsheet,
  otherwise logins fail and data writes are rejected.
* **Locked accounts.** Four failed attempts lock an account for 30 minutes (columns I and J of
  the `Users` sheet). Failed attempts from the imported copy are recorded in the shared sheet,
  so a user can be locked out for both projects. Clear columns I and J for that row.
* **Sessions are shared.** With "Execute as: Me", every visitor runs as the deployer, so the
  session store is shared. If two people sign in at once they can overwrite each other's
  session. Deploying with "Execute as: the user accessing" avoids this.
* **Owner-only features.** Test Mode and licence activation only work for the user whose
  `Users` sheet row has the system-owner address in the `googleEmail` column. That address is
  no longer hard-coded; `getSystemOwnerEmail()` resolves it in this order:

  1. the script property `OWNER_EMAIL` — set it by running
     **`setSystemOwnerEmail('you@yourdomain.com')`** in the Apps Script editor;
  2. a **System Owner Email** row in the `Settings` sheet;
  3. the `TEST_MODE_OWNER_EMAIL` constant in `Code.gs` (the original address, kept as fallback).

  The Test Mode dialog in `Index.html` now prints whatever address is in force instead of a
  hard-coded one.
