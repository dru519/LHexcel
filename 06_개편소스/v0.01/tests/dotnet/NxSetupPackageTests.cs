using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using LH.NxSetup;

internal static class NxSetupPackageTests
{
    static int passed;
    static void Expect(bool value) { if (!value) throw new Exception("Assertion failed"); passed++; }
    static void Reject(Action action)
    {
        try { action(); } catch (InvalidDataException) { passed++; return; }
        throw new Exception("Expected InvalidDataException");
    }
    static Dictionary<string, byte[]> Fixture()
    {
        var data = new Dictionary<string, byte[]>();
        foreach (string name in new[] { "README.md", "payload/Product.xlam", "payload/NxHost32.dll",
            "payload/NxHost64.dll", "payload/docs/r63-build.md", "deployment/Install-NxEnhanced.ps1",
            "deployment/Register-NxHost.ps1", "deployment/Uninstall-NxEnhanced.ps1", "deployment/Verify-NxEnhancedPackage.ps1" })
            data.Add(name, Encoding.UTF8.GetBytes("synthetic fixture; not installable: " + name));
        data.Add("payload/SHA256SUMS", Manifest(data, "payload/"));
        data.Add("PACKAGE_SHA256SUMS", Manifest(data, ""));
        return data;
    }
    static byte[] Manifest(Dictionary<string, byte[]> data, string prefix)
    {
        return Encoding.UTF8.GetBytes(String.Join("\n", data.Keys.Where(k => k.StartsWith(prefix))
            .OrderBy(k => k, StringComparer.Ordinal).Select(k => Package.Hash(data[k]) + "  " + k.Substring(prefix.Length))) + "\n");
    }
    static byte[] Zip(Dictionary<string, byte[]> files, string extra)
    {
        using (var buffer = new MemoryStream())
        {
            using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, true))
            {
                foreach (var file in files)
                    using (var target = zip.CreateEntry(file.Key).Open()) target.Write(file.Value, 0, file.Value.Length);
                if (extra != null) using (var target = zip.CreateEntry(extra).Open()) target.WriteByte(42);
            }
            return buffer.ToArray();
        }
    }
    static int Main()
    {
        try
        {
            Expect(Package.Hash(new byte[0]) == "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
            var files = Fixture();
            Expect(Package.Read(Zip(files, null), "v0.01_r63").Count == 11);
            foreach (string extra in new[] { "../escape", "/absolute", "C:/escape", "payload/../escape",
                "payload\\escape", "README.md", "readme.md", "unexpected.exe", "payload/ADS:stream", "payload/." })
                Reject(() => Package.Read(Zip(files, extra), "v0.01_r63"));
            Reject(() => Package.Read(Zip(files, null), "v0.01_r62"));
            Reject(() => Package.Read(Zip(files, null), "../r63"));
            files["payload/Product.xlam"] = new byte[] { 1 };
            Reject(() => Package.Read(Zip(files, null), "v0.01_r63"));
            files = Fixture(); files.Remove("README.md");
            Reject(() => Package.Read(Zip(files, null), "v0.01_r63"));
            files = Fixture();
            files["PACKAGE_SHA256SUMS"] = Encoding.UTF8.GetBytes("bad manifest");
            Reject(() => Package.Read(Zip(files, null), "v0.01_r63"));
            files = Fixture();
            string root = Path.Combine(Path.GetTempPath(), "nx-setup-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            foreach (var file in files)
            {
                string path = Path.Combine(root, file.Key.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllBytes(path, file.Value);
            }
            Package.VerifyDirectory(root, files); passed++;
            File.WriteAllText(Path.Combine(root, "README.md"), "changed");
            Reject(() => Package.VerifyDirectory(root, files));
            Console.WriteLine("PASS|NxSetupPackageTests|" + passed + "|SYNTHETIC_ONLY|fixture=" + root);
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
