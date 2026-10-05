using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace LH.NxSetup
{
    public static class Package
    {
        const long MaxEntry = 64L * 1024 * 1024;
        const long MaxTotal = 128L * 1024 * 1024;

        public static string Hash(byte[] bytes)
        {
            using (SHA256 sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }

        public static void SafeRelative(string name)
        {
            if (String.IsNullOrEmpty(name) || name.Contains("\\") || name.Contains(":") ||
                name.StartsWith("/") || name.EndsWith("/") || name.Split('/').Any(
                    p => p == "" || p == "." || p == ".." || p.EndsWith(".") || p.EndsWith(" ")))
                throw new InvalidDataException("Unsafe package path");
        }

        public static Dictionary<string, byte[]> Read(byte[] zip, string version)
        {
            if (!Regex.IsMatch(version, @"\Av0\.01_r[0-9]+\z"))
                throw new InvalidDataException("Invalid version");
            string release = version.Split('_')[1];
            var expected = new HashSet<string>(new[] {
                "README.md", "PACKAGE_SHA256SUMS", "payload/Product.xlam",
                "payload/NxHost32.dll", "payload/NxHost64.dll", "payload/SHA256SUMS",
                "payload/NxCore32.dll", "payload/NxCore64.dll",
                "payload/docs/" + release + "-build.md",
                "deployment/Install-NxEnhanced.ps1", "deployment/Register-NxHost.ps1",
                "deployment/Uninstall-NxEnhanced.ps1", "deployment/Verify-NxEnhancedPackage.ps1"
            }, StringComparer.Ordinal);
            var files = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            long total = 0;
            using (var stream = new MemoryStream(zip, false))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                bool wrapped = archive.Entries.Count > 0 && archive.Entries.All(e => e.FullName.StartsWith("패키지/", StringComparison.Ordinal));
                foreach (var entry in archive.Entries)
                {
                    string name = wrapped ? entry.FullName.Substring("패키지/".Length) : entry.FullName;
                    SafeRelative(name);
                    if (!expected.Contains(name) || files.ContainsKey(name) ||
                        entry.Length < 0 || entry.Length > MaxEntry || (total += entry.Length) > MaxTotal ||
                        ((entry.ExternalAttributes >> 16) & 0xF000) == 0xA000)
                        throw new InvalidDataException("Unexpected, duplicate or oversized package entry");
                    using (var source = entry.Open())
                    using (var target = new MemoryStream())
                    {
                        byte[] buffer = new byte[65536];
                        int read;
                        while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            if (target.Length + read > entry.Length)
                                throw new InvalidDataException("Expanded ZIP size mismatch");
                            target.Write(buffer, 0, read);
                        }
                        if (target.Length != entry.Length) throw new InvalidDataException("Truncated ZIP");
                        files.Add(name, target.ToArray());
                    }
                }
            }
            if (!expected.SetEquals(files.Keys)) throw new InvalidDataException("Missing package entry");
            VerifyManifest(files, "PACKAGE_SHA256SUMS", "");
            VerifyManifest(files, "payload/SHA256SUMS", "payload/");
            return files;
        }

        static void VerifyManifest(Dictionary<string, byte[]> files, string name, string prefix)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            string previous = null;
            foreach (var line in new UTF8Encoding(false, true).GetString(files[name]).Split('\n'))
            {
                if (line == "") continue;
                var match = Regex.Match(line.TrimEnd('\r'), @"\A([0-9a-f]{64})  (.+)\z");
                if (!match.Success) throw new InvalidDataException("Invalid hash manifest");
                string relative = match.Groups[2].Value;
                SafeRelative(relative);
                string path = prefix + relative;
                if (!seen.Add(path) || path == name || !files.ContainsKey(path) ||
                    (previous != null && String.CompareOrdinal(previous, relative) >= 0) ||
                    Hash(files[path]) != match.Groups[1].Value)
                    throw new InvalidDataException("Package SHA256 mismatch");
                previous = relative;
            }
            if (!seen.SetEquals(files.Keys.Where(n => n.StartsWith(prefix, StringComparison.Ordinal) && n != name)))
                throw new InvalidDataException("Hash manifest inventory mismatch");
        }

        public static void RejectReparseAncestors(string path)
        {
            for (var item = new DirectoryInfo(Path.GetFullPath(path)); item != null; item = item.Parent)
                if (item.Exists && (item.Attributes & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Reparse installation path is not allowed");
        }

        public static void VerifyDirectory(string root, Dictionary<string, byte[]> files)
        {
            RejectReparseAncestors(root);
            // Walk each level without descending into junctions.
            var pending = new Stack<string>();
            pending.Push(root);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            while (pending.Count != 0)
                foreach (string item in Directory.GetFileSystemEntries(pending.Pop()))
                {
                    FileAttributes attr = File.GetAttributes(item);
                    if ((attr & FileAttributes.ReparsePoint) != 0)
                        throw new IOException("Reparse package item");
                    if ((attr & FileAttributes.Directory) != 0) { pending.Push(item); continue; }
                    string relative = item.Substring(root.TrimEnd(Path.DirectorySeparatorChar).Length + 1).Replace('\\', '/');
                    if (!files.ContainsKey(relative) || !seen.Add(relative) ||
                        new FileInfo(item).Length != files[relative].Length ||
                        Hash(File.ReadAllBytes(item)) != Hash(files[relative]))
                        throw new InvalidDataException("Existing installer payload changed");
                }
            if (!seen.SetEquals(files.Keys)) throw new InvalidDataException("Incomplete installer payload");
        }

        public static string Stage(Dictionary<string, byte[]> files, string version, string packageHash)
        {
            string final = CachedPath(version, packageHash);
            string parent = Path.GetDirectoryName(final);
            RejectReparseAncestors(parent);
            Directory.CreateDirectory(parent);
            if (Directory.Exists(final)) { VerifyDirectory(final, files); return final; }
            string stage = Path.Combine(parent, "staging-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stage);
            try
            {
                foreach (var file in files)
                {
                    SafeRelative(file.Key);
                    string path = Path.Combine(stage, file.Key.Replace('/', Path.DirectorySeparatorChar));
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                        output.Write(file.Value, 0, file.Value.Length);
                }
                VerifyDirectory(stage, files);
                Directory.Move(stage, final);
                return final;
            }
            catch
            {
                // Preserve the exact failed staging folder for diagnosis; never recursive-delete user paths.
                throw new IOException("Installer staging failed; diagnostic folder: " + stage);
            }
        }

        public static string CachedPath(string version, string packageHash)
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LHexcel", "Setup", version, packageHash);
        }
    }
}
