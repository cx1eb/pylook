# 📬 pylook

A lightweight Python wrapper around Microsoft Outlook's COM interface (`pywin32`)
that makes automating Outlook feel as clean as using PowerShell without the
PowerShell.

It is designed for IT support, automation, and email management tasks where
you need to navigate shared mailboxes, walk folder trees, and move/copy
emails between folders **reliably**.

---

## ✨ Features

- ✅ Connect to a running Outlook instance (or start one)
- ✅ Detect existing vs newly-created Outlook sessions
- ✅ Built-in debug logging (`debug=True`)
- ✅ List shared mailboxes attached to the current profile
- ✅ Get mailbox root by display name or SMTP
- ✅ Navigate folders by path (list or string)
- ✅ Walk entire folder trees recursively
- ✅ Reliable iteration (works around Outlook COM iterator bugs)
- ✅ Count emails (top-level or recursive)
- ✅ Copy and move emails one-by-one (safe + verifiable)
- ✅ Bulk safe-move with verification, retry-friendly design
- ✅ Friendly COM error names via custom `ComError` enum

---

## 📦 Installation

Requires Python 3.9+ on Windows with Outlook installed.
Currently not its own pip package.

```bash
pip install pywin32
```

Then drop the `pylook/` package into your project.

---

## 📁 Package structure

```
pylook/
├── __init__.py        # exposes OutlookClient
├── outlook.py         # main OutlookClient class
└── enums.py           # ComError + ExchangeStoreType enums
```

Use it from your script:

```python
from pylook import OutlookClient
```

---

## ⚠️ Important: Online vs Cached Exchange Mode

If you’re working with **shared mailboxes**, you may find that some
folders are missing from COM but show up fine in OWA.

✅ The fix is to **turn off** Cached Exchange Mode for that account:

> **File → Account Settings → Account Settings → Double-click account →
> Uncheck _Use Cached Exchange Mode to download email to an Outlook data file_**

In **Online Mode**, COM talks directly to Exchange and sees the true
folder structure — exactly what OWA sees.

---

## 🚀 Quick start

```python
from pylook import OutlookClient

client = OutlookClient(debug=True)

if not client.connect():
    print("Failed to connect to Outlook.")
    exit(1)

print(f"Connected: {client.is_connected}")
print(f"Version: {client.version}")
print(f"Profile: {client.default_profile}")
print(f"Accounts: {client.account_count}")
```

---

## 📚 Method reference & examples

### 🔌 `connect()`

Connects to a running Outlook instance or starts a new one.

```python
client = OutlookClient(debug=True)
client.connect()
```

---

### 📬 `get_shared_mailboxes()`

Returns all shared / additional mailboxes attached to the profile.

```python
for mb in client.get_shared_mailboxes():
    print(f"{mb['name']} ({mb['smtp']})")
```

Output:
```
Shared Mailbox (None)
team@example.com (team@example.com)
```

---

### 🌳 `get_mailbox_root(mailbox_name)`

Returns the root folder of a mailbox by its **display name**.

```python
root = client.get_mailbox_root("Shared Mailbox")
print(root.FolderPath)
```

---

### 📥 `get_shared_inbox(smtp_address)`

Returns a shared mailbox's Inbox folder **resolved via Exchange**
(bypasses local cache issues).

```python
inbox = client.get_shared_inbox("team@example.com")
mailbox = inbox.Parent   # full mailbox root

print(inbox.Name, inbox.FolderPath)
print("Subfolder count:", inbox.Folders.Count)
```

---

### 📁 `list_folders(folder)`

Returns the **direct** subfolders of a folder (one level deep).

```python
for f in client.list_folders(inbox):
    print(f.Name)
```

---

### 🌲 `walk_folders(folder)`

Recursively walks **every folder** under a starting folder.
Yields tuples of `(folder, depth)`.

```python
for folder, depth in client.walk_folders(mailbox):
    print("  " * depth + folder.Name)
```

Sample output:
```
Shared Mailbox
  Inbox
    2024
    2025
  Sent Items
  Archive
```

---

### 🗺️ `get_folder_by_path(root, path)`

Navigate to a folder using either a **list** or **string** path.

```python
# Recommended: list of names (no ambiguity)
target = client.get_folder_by_path(root, ["Inbox", "Projects", "2025"])

# Or string with separator
target = client.get_folder_by_path(root, "Inbox/Projects/2025", sep="/")

print(target.FolderPath)
```

---

### 🔢 `count_emails(folder, include_subfolders=False)`

Counts emails in a folder. Top-level only by default.

```python
n = client.count_emails(target)
print(f"This folder has {n} emails")

total = client.count_emails(target, include_subfolders=True)
print(f"Including subfolders: {total}")
```

---

### 📨 `copy_email(item, target_folder)`

Copies a single email into another folder.

```python
src = client.get_folder_by_path(root, ["Inbox"])
dst = client.get_folder_by_path(root, ["Archive"])

first_email = src.Items.Item(1)
client.copy_email(first_email, dst)
```

---

### 🚚 `move_email(item, target_folder)`

Moves a single email into another folder.

```python
client.move_email(first_email, dst)
```

---

### 🛡️ `move_emails_safely(source, target, limit=None, delay=0, verify=True)`

Safely move emails **one at a time** with full verification.
Handles Outlook security limits and ensures no email is lost mid-run.

```python
src = client.get_folder_by_path(root, ["Inbox", "To Archive"])
dst = client.get_folder_by_path(root, ["Archive", "2024"])

stats = client.move_emails_safely(
    source_folder=src,
    target_folder=dst,
    limit=10,        # safe test: only first 10
    delay=0.2,       # avoid throttling
    verify=True,     # confirm counts after each move
)

print(stats)
```

Returns a dict like:

```python
{
    "expected": 10,
    "moved": 10,
    "failed": 0,
    "failed_subjects": [],
    "source_before": 152,
    "source_after": 142,
    "target_before": 4,
    "target_after": 14,
    "target_delta": 10,
    "verified": True,
}
```

---

### ♻️ `move_until_empty(source, target, batch_size=50, delay=0.1)`

Drain a folder in batches — perfect for huge mailboxes
or unattended migrations.

```python
client.move_until_empty(src, dst, batch_size=100, delay=0.2)
```

---

## 🧰 Properties

These behave like read-only attributes:

```python
client.is_connected     # True/False
client.version          # Outlook version string
client.name             # "Outlook"
client.account_count    # number of mail accounts
client.default_profile  # profile name
```

---

## 🐞 Debug mode

Enable debug logging with `debug=True`:

```python
client = OutlookClient(debug=True)
client.connect()
```

You'll see logs like:

```
[PYO]: Outlook already running, connecting to instance.
[PYO]: Outlook version: 16.0.0.20026
[PYO]: Accounts available: 1
[PYO]: Entered folder: Inbox
[PYO]: Entered folder: Projects
```

---

## 🧠 Common pitfalls (and how this wrapper avoids them)

| Problem | What we do |
|---------|------------|
| Outlook iterator skips folders | Use index-based `.Item(i)` access |
| Forward iteration during moves skips items | Always iterate **backwards** |
| Shared mailbox missing folders | Use `get_shared_inbox().Parent` + Online Mode |
| Cryptic COM errors | Map HRESULTs to friendly `ComError` enum |
| Outlook security blocking bulk ops | Move one-by-one with `move_emails_safely` |
| No verification | Built-in count comparison after each move |

---

## 🛣️ Roadmap

Planned next features:

- 🗂️ `copy_folder(source, target)` — copy a folder + contents
- 🌳 `mirror_folder_tree(source, target)` — recursive structure migration
- 🔍 `find_emails(folder, filter)` — filter by subject/date/sender
- 📦 `export_folder(folder, path)` — save emails to disk (.msg/.eml)
- 📊 Progress callbacks + `tqdm` integration
- 🧪 Unit tests

---

## 📝 License

MIT — do whatever you want, but no warranty.

---

## 🙏 Acknowledgements

Built on top of [`pywin32`](https://pypi.org/project/pywin32/),
the Python bindings for the Windows COM API.

This wrapper exists to make Outlook automation in Python feel less
like wrestling with a 25-year-old COM interface — and more like
writing modern Python.
