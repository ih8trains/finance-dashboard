# T-Bank Auto Sync

Chrome extension, Manifest V3. Current version: **4.0.6**.

## Install
1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Choose **Load unpacked** and select this directory.

## Use
1. Open and authorize on https://www.tbank-online.com/ or a supported T-Bank domain.
2. Open the extension popup.
3. Enter the archive password and GitHub token in **Доступ к архивам**.
4. Start synchronization.

The extension reads recent T-Bank operations and updates encrypted monthly archives in this repository. It does not use Google Sheets or Apps Script.

## Credentials
The archive password and GitHub token are stored in Chrome `storage.local` on this browser profile. They are not embedded in the source files. Do not use a shared/untrusted Chrome profile. Use a fine-grained GitHub token with access limited to this repository and the minimum required Contents permissions.

## Change in 4.0.6
Existing operations are compared by bank fields; the manual `Моя группа` column is excluded. Unchanged months are skipped completely: no new encryption, backup, archive upload, or manifest upload. A changed bank field counts as one update; a new operation counts as one addition.
