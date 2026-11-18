# Excel Studio Pro Cloud

GitHub Pages app with Google sign-in and per-user workbook sync through Firebase Realtime Database.

## Firebase setup required

Before cloud sync works, a Firebase project administrator must:

1. Enable Google under Firebase Authentication > Sign-in method.
2. Add `isbchat.github.io` to Firebase Authentication's authorized domains and to the Google OAuth web client's authorized JavaScript origins.
3. Merge `firebase-rules-fragment.json` into the existing Realtime Database rules while preserving existing messenger paths. Remove broad parent-level read/write grants; Realtime Database grants cascade and a child rule cannot restrict an already-open parent.

Workbooks are stored at `excelStudioWorkbooks/<Firebase UID>/workbook` so each signed-in user has an isolated data path.