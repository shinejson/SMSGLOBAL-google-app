# SMSGLOBAL-google-app
post

## Login breaks on an imported copy of the project?

If the app signs in fine on the account where it was built but not on another Gmail account
bound to the same spreadsheet, read **[CROSS_ACCOUNT_LOGIN_FIX.md](CROSS_ACCOUNT_LOGIN_FIX.md)**.

Short version: the password salt used to live in Script Properties, which belong to the Apps
Script *project* and are **not** copied when you import the project into another account. The
copy generated its own salt, so every stored password hash stopped matching. The salt is now
stored in the spreadsheet (hidden `_SystemConfig` sheet) and shared by every copy.
