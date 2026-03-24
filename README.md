# 📧 PST Exporter

Export every email from a Microsoft Outlook `.pst` file into individual `.msg` files.

This utility uses Outlook on Windows to open the PST, walk the folder tree, and save each mail item as a standalone message file.

## ✨ Features

- Export emails from a `.pst` file to `.msg`
- Preserve the original PST folder structure
- Handle duplicate subject lines automatically
- Skip non-mail Outlook items
- Sanitize filenames for Windows compatibility
- Work with password-protected PSTs through Outlook's normal prompt

### Duplicate subject handling

If multiple emails in the same folder have the same subject, filenames are de-duplicated automatically:

- `Quarterly Update.msg`
- `Quarterly Update (1).msg`
- `Quarterly Update (2).msg`

## 🧰 Requirements

- Windows 11
- Microsoft Outlook (Classic) installed locally
- .NET 8 SDK or .NET 8 runtime

## 🚀 Build

```powershell
dotnet build PSTExporterApp.csproj -c Release
```

## ▶️ Run

### Using `dotnet run`

```powershell
dotnet run --project PSTExporterApp.csproj -- "C:\Mail\archive.pst" "C:\Exports\archive"
```

### Using the built executable

```powershell
.\bin\Release\net8.0-windows\PSTExporterApp.exe "C:\Mail\archive.pst" "C:\Exports\archive"
```

## 🔐 Password-protected PSTs

If the PST is password protected, Outlook may display its normal password prompt when the file is opened.

- Enter the correct password in the Outlook prompt
- If you do not see the prompt immediately, check behind other windows
- Once the PST opens successfully, export continues as normal

The app does not accept a password as a command-line argument. Password entry is handled by Outlook itself.

## 🗂️ Output behavior

- A folder is created for each PST folder
- Only mail items are exported to `.msg`
- Non-mail items are skipped
- Invalid Windows filename characters are replaced
- Reserved Windows device names are adjusted automatically

## 📝 Notes

- Outlook must be installed and available on the machine running the exporter
- The app starts Outlook through COM automation
- Very large PSTs may take some time to process depending on Outlook performance
