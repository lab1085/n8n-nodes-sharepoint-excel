# SharePoint File Locking Behavior

This document explains SharePoint file locking behavior and how it affects the n8n SharePoint Excel node.

## The Error

When attempting to write to a file that is open by another user (or yourself on another device), you will see:

```
Error: File is locked - someone has it open in Excel

Description: The file cannot be modified because it is open in Excel or SharePoint.
Close the file and try again. If using a shared file, wait for other users to close it.
```

## Why Locks Persist After Closing

SharePoint doesn't release locks immediately after closing a file. This is by design.

| Cause                         | Why                                                            |
| ----------------------------- | -------------------------------------------------------------- |
| **Co-authoring grace period** | SharePoint waits in case you reconnect                         |
| **Upload Center**             | Windows may still have the file "open" in Office Upload Center |
| **Auto-save sync**            | Waits to ensure all changes are synced before releasing        |
| **Browser session**           | Excel Online session may not signal closure properly           |
| **Hung Excel process**        | Desktop Excel didn't fully close (running in background)       |

## Lock Timeout Duration

Microsoft doesn't document an exact timeout, but based on community reports:

- **~10-15 minutes** is the typical timeout for auto-release
- Sometimes longer if Upload Center is stuck
- In rare cases, locks can persist for hours if there's a sync issue

## Quick Fixes for Users

### 1. Check Office Upload Center (Windows)

1. Look in the system tray for the Office Upload Center icon
2. Open it and check for stuck uploads
3. Clear any pending uploads for the file

### 2. Close Excel Fully

1. Open Task Manager (Ctrl+Shift+Esc)
2. Look for any running Excel processes
3. End all Excel processes

### 3. Sign Out of Office Online

1. Go to office.com
2. Sign out completely
3. This forces the session cleanup

### 4. Wait

SharePoint eventually releases the lock automatically (10-15 minutes typical).

### 5. Check Other Devices

The file may be open on:

- Another computer
- Mobile device with Office app
- Browser tab you forgot about

## Recommendations for Workflows

### Use Dedicated Files

Don't share automation files with users who manually edit them. Create separate files for:

- Manual user editing
- Automated workflows

### Schedule Off-Hours

Run workflows when no one is editing:

- Before business hours
- After business hours
- Weekends

### Add Retry Logic in n8n

Use n8n's built-in retry functionality:

1. Open node settings
2. Go to "Settings" tab
3. Enable "Retry On Fail"
4. Set retry attempts (e.g., 3)
5. Set wait between retries (e.g., 30 seconds)

### Use Error Handling

Add an error workflow that:

1. Catches the locked file error
2. Waits a set time
3. Retries the operation
4. Notifies the team if it still fails

## Technical Details

### How the Node Works

The node uses a download-modify-upload pattern:

```
GET /content  →  modify locally  →  PUT /content
                                       ↑
                              Fails if file is locked
```

SharePoint rejects the `PUT` request because another session has the file open.

### SharePoint Lock Types

| Lock Type          | Duration      | Cause                       |
| ------------------ | ------------- | --------------------------- |
| **Exclusive**      | While open    | File open in desktop Excel  |
| **Shared/Co-auth** | While open    | File open in Excel Online   |
| **Orphaned**       | 10-15 min     | Session closed improperly   |
| **Upload Center**  | Until cleared | Sync stuck in Upload Center |

### Error Codes

The node detects these error patterns:

- HTTP status `423` (Locked)
- Error code `resourceLocked`
- Error code `notAllowed` with locked message

## References

- [Excel file is locked for editing - Microsoft Support](https://support.microsoft.com/en-us/office/excel-file-is-locked-for-editing-6fa93887-2c2c-45f0-abcc-31b04aed68b3)
- [File Locking in SharePoint Online - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5095587/file-locking-in-sharepoint-online)
- [SharePoint File Locks - Power Platform Community](https://powerusers.microsoft.com/t5/General-Power-Automate/SharePoint-File-Locks/td-p/2184083)
