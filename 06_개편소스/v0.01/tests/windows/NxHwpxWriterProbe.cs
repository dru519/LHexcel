using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Text;
using System.Threading;
internal static class NxHwpxWriterProbe {
    static int Main(string[] args) {
        try {
            if (args.Length == 2 && args[0] == "fixtures") { CreateFixtures(args[1]); Console.WriteLine("PASS|HwpxFixtures"); return 0; }
            var type = Assembly.LoadFrom(args[0]).GetType("LH.NxHost.HwpxTableWriter", true);
            var write = type.GetMethod("Write", BindingFlags.Static | BindingFlags.NonPublic);
            var cancel = new CancellationTokenSource();
            if (args.Length > 4 && args[4] == "cancel") cancel.Cancel();
            write.Invoke(null, new object[] { File.ReadAllText(args[1]), args[2], args[3], cancel.Token });
            Console.WriteLine("PASS|HwpxWriter"); return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    static void CreateFixtures(string folder) {
        Directory.CreateDirectory(folder);
        using (var bitmap = new Bitmap(40, 30)) { using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.Red); bitmap.Save(Path.Combine(folder, "C missing.png"), ImageFormat.Png); }
        CreateExifJpeg(Path.Combine(folder, "B late.jpg"), Color.Blue, "2024:01:02 03:04:05", 1);
        CreateExifJpeg(Path.Combine(folder, "A early.jpg"), Color.Green, "2020:01:02 03:04:05", 6);
    }
    static void CreateExifJpeg(string path, Color color, string taken, ushort orientation) {
        byte[] jpeg;
        using (var bitmap = new Bitmap(20, 30)) { using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(color); using (var memory = new MemoryStream()) { bitmap.Save(memory, ImageFormat.Jpeg); jpeg = memory.ToArray(); } }
        byte[] stamp = Encoding.ASCII.GetBytes(taken + "\0");
        byte[] block;
        using (var memory = new MemoryStream()) using (var writer = new BinaryWriter(memory)) {
            writer.Write(Encoding.ASCII.GetBytes("Exif\0\0")); writer.Write((byte)'I'); writer.Write((byte)'I'); writer.Write((ushort)42); writer.Write((uint)8);
            writer.Write((ushort)2); writer.Write((ushort)0x0112); writer.Write((ushort)3); writer.Write((uint)1); writer.Write(orientation); writer.Write((ushort)0);
            writer.Write((ushort)0x8769); writer.Write((ushort)4); writer.Write((uint)1); writer.Write((uint)38); writer.Write((uint)0);
            writer.Write((ushort)1); writer.Write((ushort)0x9003); writer.Write((ushort)2); writer.Write((uint)stamp.Length); writer.Write((uint)56); writer.Write((uint)0); writer.Write(stamp); block = memory.ToArray();
        }
        using (var output = new FileStream(path, FileMode.CreateNew)) using (var writer = new BinaryWriter(output)) {
            writer.Write(jpeg, 0, 2); writer.Write((byte)0xff); writer.Write((byte)0xe1); ushort length = (ushort)(block.Length + 2); writer.Write((byte)(length >> 8)); writer.Write((byte)length); writer.Write(block); writer.Write(jpeg, 2, jpeg.Length - 2);
        }
    }
}
