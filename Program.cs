using System.Runtime.InteropServices;

namespace PSTExporterApp;

internal static class Program
{
    private const string DefaultSubject = "No Subject";
    private const int OlObjectClassMail = 43;
    private const int OlSaveAsTypeMsgUnicode = 9;

    [STAThread]
    private static int Main(string[] args)
    {
        if (HasHelpArgument(args))
        {
            PrintUsage();
            return 0;
        }

        if (args.Length != 2)
        {
            PrintUsage();
            return 1;
        }

        var pstPath = Path.GetFullPath(args[0]);
        var outputPath = Path.GetFullPath(args[1]);

        if (!File.Exists(pstPath))
        {
            Console.Error.WriteLine($"PST file not found: {pstPath}");
            return 1;
        }

        Directory.CreateDirectory(outputPath);

        object? application = null;
        object? session = null;
        object? store = null;
        object? rootFolder = null;

        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Microsoft Outlook is not installed or is not registered correctly.");

            application = Activator.CreateInstance(outlookType)
                ?? throw new InvalidOperationException("Unable to start Microsoft Outlook.");

            session = application.GetType().InvokeMember("Session", System.Reflection.BindingFlags.GetProperty, null, application, null)
                ?? throw new InvalidOperationException("Unable to access the Outlook MAPI session.");

            Console.WriteLine("Opening PST in Outlook...");
            Console.WriteLine("If the PST is password protected, Outlook may show a password prompt.");
            Console.WriteLine("If you do not see it, check behind other windows and complete it to continue.");
            Console.WriteLine();

            session.GetType().InvokeMember("AddStore", System.Reflection.BindingFlags.InvokeMethod, null, session, [pstPath]);
            store = FindStore(session, pstPath)
                ?? throw new InvalidOperationException($"Unable to open PST store: {pstPath}");

            rootFolder = store.GetType().InvokeMember("GetRootFolder", System.Reflection.BindingFlags.InvokeMethod, null, store, null)
                ?? throw new InvalidOperationException("Unable to read the root folder from the PST.");

            var exporter = new MailExporter();
            exporter.ExportFolder(rootFolder, outputPath);

            Console.WriteLine($"Exported {exporter.ExportedCount} mail item(s) to {outputPath}");
            if (exporter.SkippedItemCount > 0)
            {
                Console.WriteLine($"Skipped {exporter.SkippedItemCount} non-mail item(s).");
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
        finally
        {
            if (session is not null && rootFolder is not null)
            {
                try
                {
                    session.GetType().InvokeMember("RemoveStore", System.Reflection.BindingFlags.InvokeMethod, null, session, [rootFolder]);
                }
                catch
                {
                }
            }

            ReleaseComObject(rootFolder);
            ReleaseComObject(store);
            ReleaseComObject(session);
            ReleaseComObject(application);
        }
    }

    private static bool HasHelpArgument(IEnumerable<string> args) =>
        args.Any(arg => string.Equals(arg, "--help", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(arg, "-h", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(arg, "/?", StringComparison.OrdinalIgnoreCase));

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  PSTExporterApp.exe <path-to-pst> <output-folder>");
        Console.WriteLine();
        Console.WriteLine("Example:");
        Console.WriteLine(@"  PSTExporterApp.exe C:\Mail\archive.pst C:\Exports\archive");
        Console.WriteLine();
        Console.WriteLine("Requirements:");
        Console.WriteLine("  Microsoft Outlook must be installed on this machine.");
    }

    private static object? FindStore(object session, string pstPath)
    {
        var normalizedPstPath = Path.GetFullPath(pstPath);
        object? stores = null;

        try
        {
            stores = session.GetType().InvokeMember("Stores", System.Reflection.BindingFlags.GetProperty, null, session, null);
            var count = Convert.ToInt32(stores!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, stores, null));

            for (var i = 1; i <= count; i++)
            {
                object? candidate = null;
                try
                {
                    candidate = stores.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, stores, [i]);
                    var filePath = candidate?.GetType().InvokeMember("FilePath", System.Reflection.BindingFlags.GetProperty, null, candidate, null) as string;
                    if (candidate is not null &&
                        !string.IsNullOrWhiteSpace(filePath) &&
                        string.Equals(Path.GetFullPath(filePath), normalizedPstPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return candidate;
                    }
                }
                catch
                {
                    ReleaseComObject(candidate);
                    throw;
                }

                ReleaseComObject(candidate);
            }
        }
        finally
        {
            ReleaseComObject(stores);
        }

        return null;
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }

    private sealed class MailExporter
    {
        public int ExportedCount { get; private set; }

        public int SkippedItemCount { get; private set; }

        public void ExportFolder(object folder, string parentOutputPath)
        {
            var folderName = folder.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, folder, null) as string;
            var currentOutputPath = CreateFolderPath(parentOutputPath, folderName ?? "Root");

            ExportItems(folder, currentOutputPath);
            ExportSubfolders(folder, currentOutputPath);
        }

        private void ExportItems(object folder, string outputPath)
        {
            object? items = null;

            try
            {
                items = folder.GetType().InvokeMember("Items", System.Reflection.BindingFlags.GetProperty, null, folder, null);
                var count = Convert.ToInt32(items!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, items, null));

                for (var i = 1; i <= count; i++)
                {
                    object? item = null;
                    try
                    {
                        item = items.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, items, [i]);
                        var objectClass = Convert.ToInt32(item?.GetType().InvokeMember("Class", System.Reflection.BindingFlags.GetProperty, null, item, null));

                        if (objectClass == OlObjectClassMail)
                        {
                            ExportMailItem(item!, outputPath);
                        }
                        else
                        {
                            SkippedItemCount++;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(item);
                    }
                }
            }
            finally
            {
                ReleaseComObject(items);
            }
        }

        private void ExportSubfolders(object folder, string outputPath)
        {
            object? folders = null;

            try
            {
                folders = folder.GetType().InvokeMember("Folders", System.Reflection.BindingFlags.GetProperty, null, folder, null);
                var count = Convert.ToInt32(folders!.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, folders, null));

                for (var i = 1; i <= count; i++)
                {
                    object? childFolder = null;
                    try
                    {
                        childFolder = folders.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, folders, [i]);
                        if (childFolder is not null)
                        {
                            ExportFolder(childFolder, outputPath);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(childFolder);
                    }
                }
            }
            finally
            {
                ReleaseComObject(folders);
            }
        }

        private void ExportMailItem(object mailItem, string outputPath)
        {
            var subject = mailItem.GetType().InvokeMember("Subject", System.Reflection.BindingFlags.GetProperty, null, mailItem, null) as string;
            var baseFileName = SanitizeName(string.IsNullOrWhiteSpace(subject) ? DefaultSubject : subject);
            var filePath = GetUniquePath(outputPath, baseFileName, ".msg");

            mailItem.GetType().InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, null, mailItem, [filePath, OlSaveAsTypeMsgUnicode]);
            ExportedCount++;
        }

        private static string CreateFolderPath(string parentOutputPath, string folderName)
        {
            var safeFolderName = SanitizeName(string.IsNullOrWhiteSpace(folderName) ? "Root" : folderName);
            var folderPath = GetUniqueDirectoryPath(parentOutputPath, safeFolderName);
            Directory.CreateDirectory(folderPath);
            return folderPath;
        }

        private static string GetUniqueDirectoryPath(string parentPath, string baseName)
        {
            var candidate = Path.Combine(parentPath, baseName);
            if (!Directory.Exists(candidate))
            {
                return candidate;
            }

            for (var index = 1; ; index++)
            {
                candidate = Path.Combine(parentPath, $"{baseName} ({index})");
                if (!Directory.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        private static string GetUniquePath(string directoryPath, string baseName, string extension)
        {
            var candidate = Path.Combine(directoryPath, $"{baseName}{extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            for (var index = 1; ; index++)
            {
                candidate = Path.Combine(directoryPath, $"{baseName} ({index}){extension}");
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        private static string SanitizeName(string value)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitizedChars = value
                .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
                .ToArray();

            var sanitized = new string(sanitizedChars).Trim().TrimEnd('.', ' ');

            if (string.IsNullOrWhiteSpace(sanitized))
            {
                return DefaultSubject;
            }

            if (IsReservedWindowsName(sanitized))
            {
                sanitized = $"{sanitized}_";
            }

            return sanitized.Length <= 120 ? sanitized : sanitized[..120].TrimEnd();
        }

        private static bool IsReservedWindowsName(string value)
        {
            var reservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "CON",
                "PRN",
                "AUX",
                "NUL",
                "COM1",
                "COM2",
                "COM3",
                "COM4",
                "COM5",
                "COM6",
                "COM7",
                "COM8",
                "COM9",
                "LPT1",
                "LPT2",
                "LPT3",
                "LPT4",
                "LPT5",
                "LPT6",
                "LPT7",
                "LPT8",
                "LPT9"
            };

            return reservedNames.Contains(value);
        }
    }
}
