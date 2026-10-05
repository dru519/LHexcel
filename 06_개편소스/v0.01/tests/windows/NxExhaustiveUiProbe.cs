using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using Accessibility;

public sealed class NxObservedWindow
{
    public long Handle { get; set; }
    public string Title { get; set; }
    public string ClassName { get; set; }
    public string[] ChildTexts { get; set; }
    public string[] ChildClasses { get; set; }
}

public static class NxExhaustiveUiProbe
{
    private const uint BM_CLICK = 0x00F5;
    private const uint WM_CLOSE = 0x0010;
    private const uint WM_KEYDOWN = 0x0100;
    private const uint WM_KEYUP = 0x0101;
    private const uint WM_COMMAND = 0x0111;
    private const uint WM_SETTEXT = 0x000C;
    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const int IDOK = 1;
    private const int VK_RETURN = 0x0D;
    private const int VK_ESCAPE = 0x1B;
    private const uint CF_UNICODETEXT = 13;
    private const uint GMEM_MOVEABLE = 0x0002;
    private const uint OBJID_CLIENT = 0xFFFFFFFC;
    private const int ROLE_SYSTEM_TEXT = 42;
    private const int ROLE_SYSTEM_PUSHBUTTON = 43;
    private const int STATE_SYSTEM_UNAVAILABLE = 0x00000001;
    private static readonly Guid IAccessibleGuid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");

    private sealed class ProbeState
    {
        public int ProcessId;
        public IntPtr MainWindow;
        public string Mode = "popup";
        public string InputText = "native-test";
        public string OutputPath = String.Empty;
        public volatile bool StopRequested;
        public volatile bool PlanExecuteRequested;
        public volatile bool ResultCloseRequested;
        public volatile bool FileOutputInputRequested;
        public volatile bool FileOutputPreviewRequested;
        public volatile bool FileOutputSubmitRequested;
        public volatile bool FileOutputResultClosed;
        public volatile bool FileOutputFailureObserved;
        public string Failure = String.Empty;
        public readonly object Gate = new object();
        public readonly List<NxObservedWindow> Windows = new List<NxObservedWindow>();
        public readonly HashSet<string> Seen = new HashSet<string>(StringComparer.Ordinal);
        public readonly List<Thread> ActionThreads = new List<Thread>();
        public Thread Thread;
    }

    private static readonly ConcurrentDictionary<string, ProbeState> Probes =
        new ConcurrentDictionary<string, ProbeState>(StringComparer.Ordinal);

    private delegate bool EnumWindowsProc(IntPtr window, IntPtr parameter);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeInput
    {
        public uint Type;
        public NativeInputUnion Union;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct NativeInputUnion
    {
        [FieldOffset(0)] public NativeKeyboardInput Keyboard;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeKeyboardInput
    {
        public ushort VirtualKey;
        public ushort ScanCode;
        public uint Flags;
        public uint Time;
        public UIntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct GuiThreadInfo
    {
        public uint Size;
        public uint Flags;
        public IntPtr Active;
        public IntPtr Focus;
        public IntPtr Capture;
        public IntPtr MenuOwner;
        public IntPtr MoveSize;
        public IntPtr Caret;
        public NativeRect CaretRect;
    }

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(IntPtr window);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr window, StringBuilder text, int maximum);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder text, int maximum);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, string lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo information);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint count, NativeInput[] inputs, int size);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr window,
        uint objectId,
        ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out object accessibleObject);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr owner);

    [DllImport("user32.dll")]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll")]
    private static extern IntPtr GetClipboardData(uint format);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EmptyClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetClipboardData(uint format, IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll")]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalFree(IntPtr memory);

    public static string Start(int processId, long mainWindow, string mode, string inputText, string outputPath)
    {
        if (processId <= 0 || mainWindow == 0) throw new ArgumentException("Exact Excel ownership is required");
        var token = Guid.NewGuid().ToString("N");
        var state = new ProbeState
        {
            ProcessId = processId,
            MainWindow = new IntPtr(mainWindow),
            Mode = String.IsNullOrWhiteSpace(mode) ? "popup" : mode,
            InputText = String.IsNullOrEmpty(inputText) ? "native-test" : inputText,
            OutputPath = outputPath ?? String.Empty
        };
        state.Thread = new Thread(() => Run(state)) { IsBackground = true, Name = "NxExhaustiveUiProbe" };
        if (!Probes.TryAdd(token, state)) throw new InvalidOperationException("Probe token collision");
        state.Thread.Start();
        return token;
    }

    public static NxObservedWindow[] Stop(string token, int joinMilliseconds)
    {
        ProbeState state;
        if (!Probes.TryGetValue(token, out state)) return new NxObservedWindow[0];
        state.StopRequested = true;
        if (state.Thread != null && state.Thread.IsAlive) state.Thread.Join(Math.Max(1, joinMilliseconds));
        var quiesced = WaitForActionThreads(state, Math.Max(10000, joinMilliseconds));
        if (!quiesced)
        {
            CloseOwnedProbeForms(state);
            quiesced = WaitForActionThreads(state, 5000);
        }
        ProbeState removed;
        Probes.TryRemove(token, out removed);
        if (!quiesced) throw new InvalidOperationException("UI probe action threads did not stop");
        lock (state.Gate) return state.Windows.ToArray();
    }

    private static bool WaitForActionThreads(ProbeState state, int timeoutMilliseconds)
    {
        var deadline = unchecked(Environment.TickCount + Math.Max(1, timeoutMilliseconds));
        while (true)
        {
            Thread[] threads;
            lock (state.Gate)
            {
                state.ActionThreads.RemoveAll(thread => thread == null || !thread.IsAlive);
                if (state.ActionThreads.Count == 0) return true;
                threads = state.ActionThreads.ToArray();
            }
            foreach (var thread in threads)
            {
                if (thread != null && thread.IsAlive) thread.Join(50);
            }
            if (unchecked(deadline - Environment.TickCount) <= 0) return false;
        }
    }

    public static int ObservedCount(string token)
    {
        ProbeState state;
        if (!Probes.TryGetValue(token, out state)) return 0;
        lock (state.Gate) return state.Windows.Count;
    }

    public static string FailureMessage(string token)
    {
        ProbeState state;
        if (!Probes.TryGetValue(token, out state)) return String.Empty;
        lock (state.Gate) return state.Failure ?? String.Empty;
    }

    public static int ObservedTitleCount(string token, string[] expectedTitles)
    {
        ProbeState state;
        if (!Probes.TryGetValue(token, out state) || expectedTitles == null || expectedTitles.Length == 0) return 0;
        var count = 0;
        lock (state.Gate)
        {
            foreach (var window in state.Windows)
            {
                foreach (var expectedTitle in expectedTitles)
                {
                    if (String.Equals(window.Title, expectedTitle, StringComparison.Ordinal))
                    {
                        count++;
                        break;
                    }
                }
            }
        }
        return count;
    }

    public static string ReadClipboardText()
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            if (!OpenClipboard(IntPtr.Zero)) { Thread.Sleep(25); continue; }
            try
            {
                var handle = GetClipboardData(CF_UNICODETEXT);
                if (handle == IntPtr.Zero) return String.Empty;
                var pointer = GlobalLock(handle);
                if (pointer == IntPtr.Zero) return String.Empty;
                try { return Marshal.PtrToStringUni(pointer) ?? String.Empty; }
                finally { GlobalUnlock(handle); }
            }
            finally { CloseClipboard(); }
        }
        throw new InvalidOperationException("Clipboard is busy");
    }

    public static void WriteClipboardText(string value)
    {
        value = value ?? String.Empty;
        var bytes = (value.Length + 1) * 2;
        var memory = GlobalAlloc(GMEM_MOVEABLE, new UIntPtr((uint)bytes));
        if (memory == IntPtr.Zero) throw new InvalidOperationException("Clipboard allocation failed");
        var pointer = GlobalLock(memory);
        if (pointer == IntPtr.Zero) { GlobalFree(memory); throw new InvalidOperationException("Clipboard lock failed"); }
        try { Marshal.Copy((value + "\0").ToCharArray(), 0, pointer, value.Length + 1); }
        finally { GlobalUnlock(memory); }

        var transferred = false;
        try
        {
            for (var attempt = 0; attempt < 20; attempt++)
            {
                if (!OpenClipboard(IntPtr.Zero)) { Thread.Sleep(25); continue; }
                try
                {
                    if (!EmptyClipboard()) throw new InvalidOperationException("Clipboard clear failed");
                    if (SetClipboardData(CF_UNICODETEXT, memory) == IntPtr.Zero)
                        throw new InvalidOperationException("Clipboard write failed");
                    transferred = true;
                    return;
                }
                finally { CloseClipboard(); }
            }
            throw new InvalidOperationException("Clipboard is busy");
        }
        finally { if (!transferred) GlobalFree(memory); }
    }

    private static void Run(ProbeState state)
    {
        var started = Environment.TickCount;
        var lastEscapeSent = started;
        while (!state.StopRequested && unchecked(Environment.TickCount - started) < 45000)
        {
            try
            {
                var elapsed = unchecked(Environment.TickCount - started);
                if (String.Equals(state.Mode, "print-preview", StringComparison.Ordinal) && elapsed > 900 &&
                    unchecked(Environment.TickCount - lastEscapeSent) >= 500)
                {
                    SendEscapeToOwnedExcelWindows(state);
                    lastEscapeSent = Environment.TickCount;
                }
                EnumWindows((window, parameter) =>
                {
                    uint processId;
                    GetWindowThreadProcessId(window, out processId);
                    if (processId != (uint)state.ProcessId || !IsWindowVisible(window)) return true;
                    if (String.Equals(ClassName(window), "XLMAIN", StringComparison.OrdinalIgnoreCase))
                    {
                        ObserveOwnedChildForms(state, window);
                        return true;
                    }
                    HandleWindow(state, window);
                    return true;
                }, IntPtr.Zero);
                if (String.Equals(state.Mode, "navigator", StringComparison.Ordinal))
                    ObserveNavigatorTaskPane(state);
            }
            catch (Exception error)
            {
                lock (state.Gate)
                {
                    if (String.IsNullOrWhiteSpace(state.Failure))
                        state.Failure = error.GetType().Name + ": " + error.Message;
                }
                CloseOwnedProbeForms(state);
                state.StopRequested = true;
            }
            Thread.Sleep(80);
        }
    }

    private static void CloseOwnedProbeForms(ProbeState state)
    {
        try
        {
            EnumWindows((window, parameter) =>
            {
                uint processId;
                GetWindowThreadProcessId(window, out processId);
                if (processId != (uint)state.ProcessId) return true;
                if (ClassName(window).StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase))
                    PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                if (String.Equals(ClassName(window), "XLMAIN", StringComparison.OrdinalIgnoreCase))
                {
                    EnumChildWindows(window, (child, childParameter) =>
                    {
                        uint childProcessId;
                        GetWindowThreadProcessId(child, out childProcessId);
                        if (childProcessId == (uint)state.ProcessId &&
                            ClassName(child).StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase))
                            PostMessage(child, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                        return true;
                    }, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
        }
        catch { }
    }

    private static void ObserveOwnedChildForms(ProbeState state, IntPtr excelWindow)
    {
        EnumChildWindows(excelWindow, (window, parameter) =>
        {
            uint processId;
            GetWindowThreadProcessId(window, out processId);
            if (processId != (uint)state.ProcessId || !IsWindowVisible(window) ||
                !ClassName(window).StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase)) return true;
            HandleWindow(state, window);
            return true;
        }, IntPtr.Zero);
    }

    private static void SendEscapeToOwnedExcelWindows(ProbeState state)
    {
        SendEscapeToExcelWindow(state.MainWindow, state.ProcessId);
        EnumWindows((window, parameter) =>
        {
            uint processId;
            GetWindowThreadProcessId(window, out processId);
            if (window == state.MainWindow || processId != (uint)state.ProcessId ||
                !String.Equals(ClassName(window), "XLMAIN", StringComparison.OrdinalIgnoreCase)) return true;
            SendEscapeToExcelWindow(window, state.ProcessId);
            return true;
        }, IntPtr.Zero);
    }

    private static void SendEscapeToExcelWindow(IntPtr window, int expectedProcessId)
    {
        if (window == IntPtr.Zero) return;
        uint processId;
        GetWindowThreadProcessId(window, out processId);
        if (processId != (uint)expectedProcessId ||
            !String.Equals(ClassName(window), "XLMAIN", StringComparison.OrdinalIgnoreCase)) return;
        PostMessage(window, WM_KEYDOWN, new IntPtr(VK_ESCAPE), IntPtr.Zero);
        PostMessage(window, WM_KEYUP, new IntPtr(VK_ESCAPE), IntPtr.Zero);
        if (!IsWindowVisible(window)) return;
        SetForegroundWindow(window);
        Thread.Sleep(50);
        SendKeyboardKey(VK_ESCAPE);
    }

    private static void SendKeyboardKey(int virtualKey)
    {
        var inputs = new[]
        {
            new NativeInput { Type = INPUT_KEYBOARD, Union = new NativeInputUnion { Keyboard = new NativeKeyboardInput { VirtualKey = (ushort)virtualKey } } },
            new NativeInput { Type = INPUT_KEYBOARD, Union = new NativeInputUnion { Keyboard = new NativeKeyboardInput { VirtualKey = (ushort)virtualKey, Flags = KEYEVENTF_KEYUP } } }
        };
        // PostMessage above is the deterministic close path. SendInput is only
        // a best-effort fallback because Windows can reject foreground input
        // from an automation host; that rejection must never terminate the
        // verifier's background thread or its PowerShell process.
        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(NativeInput)));
    }

    private static void ReleaseMsFormsAccessible(object value)
    {
        if (value == null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); }
        catch { }
    }

    private static IntPtr FindMsFormsServer(IntPtr window)
    {
        foreach (var child in ChildWindows(window))
        {
            if (ClassName(child).StartsWith("F3 Server", StringComparison.OrdinalIgnoreCase)) return child;
        }
        throw new InvalidOperationException("Owned MSForms server is unavailable");
    }

    private static IAccessible GetMsFormsAccessibleRoot(IntPtr window)
    {
        var started = Environment.TickCount;
        while (IsWindowVisible(window) && unchecked(Environment.TickCount - started) < 1500)
        {
            try
            {
                object value;
                var interfaceId = IAccessibleGuid;
                var result = AccessibleObjectFromWindow(FindMsFormsServer(window), OBJID_CLIENT, ref interfaceId, out value);
                var accessible = value as IAccessible;
                if (result == 0 && accessible != null) return accessible;
            }
            catch (InvalidOperationException) { }
            Thread.Sleep(50);
        }
        throw new InvalidOperationException("Owned MSForms accessibility root is unavailable");
    }

    private static object FindMsFormsChildId(IAccessible accessible, string name, int role, int occurrence)
    {
        if (accessible == null) throw new InvalidOperationException("MSForms accessibility root is required");
        for (var index = 1; index <= accessible.accChildCount; index++)
        {
            object childId = index;
            string childName;
            int childRole;
            try
            {
                childName = accessible.get_accName(childId) ?? String.Empty;
                childRole = Convert.ToInt32(accessible.get_accRole(childId));
            }
            catch { continue; }
            if (role > 0 && childRole != role) continue;
            if (!String.IsNullOrEmpty(name) && !String.Equals(childName, name, StringComparison.Ordinal)) continue;
            if (occurrence-- == 0) return childId;
        }
        throw new InvalidOperationException("MSForms control is unavailable: " + (name ?? String.Empty));
    }

    private static IAccessible GetMsFormsChild(IAccessible accessible, object childId)
    {
        IAccessible child;
        try { child = accessible.get_accChild(childId) as IAccessible; }
        catch (Exception error) { throw new InvalidOperationException("MSForms child accessibility is unavailable", error); }
        if (child == null) throw new InvalidOperationException("MSForms child accessibility is unavailable");
        return child;
    }

    private static void SetMsFormsValue(IntPtr window, int textIndex, string value)
    {
        IAccessible accessible = null;
        IAccessible child = null;
        try
        {
            accessible = GetMsFormsAccessibleRoot(window);
            var childId = FindMsFormsChildId(accessible, String.Empty, ROLE_SYSTEM_TEXT, textIndex);
            child = GetMsFormsChild(accessible, childId);
            child.set_accValue(0, value ?? String.Empty);
            var readBack = child.get_accValue(0) ?? String.Empty;
            if (!String.Equals(readBack, value ?? String.Empty, StringComparison.Ordinal))
                throw new InvalidOperationException("MSForms value readback mismatch");
        }
        finally
        {
            ReleaseMsFormsAccessible(child);
            if (!Object.ReferenceEquals(child, accessible)) ReleaseMsFormsAccessible(accessible);
        }
    }

    private static bool QueueMsFormsDefaultAction(ProbeState state, IntPtr window, string name)
    {
        IAccessible accessible = null;
        try
        {
            accessible = GetMsFormsAccessibleRoot(window);
            var childId = FindMsFormsChildId(accessible, name, ROLE_SYSTEM_PUSHBUTTON, 0);
            var controlState = Convert.ToInt32(accessible.get_accState(childId));
            if ((controlState & STATE_SYSTEM_UNAVAILABLE) != 0) return false;
        }
        finally { ReleaseMsFormsAccessible(accessible); }
        Thread actionThread = null;
        actionThread = new Thread(new ThreadStart(delegate
        {
            IAccessible actionRoot = null;
            IAccessible actionChild = null;
            try
            {
                actionRoot = GetMsFormsAccessibleRoot(window);
                var actionChildId = FindMsFormsChildId(actionRoot, name, ROLE_SYSTEM_PUSHBUTTON, 0);
                actionChild = GetMsFormsChild(actionRoot, actionChildId);
                actionChild.accDoDefaultAction(0);
            }
            catch (Exception error)
            {
                if (!state.StopRequested)
                {
                    lock (state.Gate)
                    {
                        if (String.IsNullOrWhiteSpace(state.Failure))
                            state.Failure = "MSForms action " + name + ": " + error.GetType().Name + ": " + error.Message;
                    }
                    state.StopRequested = true;
                    CloseOwnedProbeForms(state);
                }
            }
            finally
            {
                ReleaseMsFormsAccessible(actionChild);
                if (!Object.ReferenceEquals(actionChild, actionRoot)) ReleaseMsFormsAccessible(actionRoot);
                lock (state.Gate) state.ActionThreads.Remove(actionThread);
            }
        })) { IsBackground = true, Name = "NxMsFormsAction-" + name };
        lock (state.Gate)
        {
            if (state.StopRequested) return false;
            state.ActionThreads.Add(actionThread);
        }
        actionThread.Start();
        return true;
    }

    private static bool IsMsFormsFileOutput(string title)
    {
        return String.Equals(title, "내엑셀 - 저장", StringComparison.Ordinal) ||
            String.Equals(title, "내엑셀 - PDF", StringComparison.Ordinal);
    }

    private static void SetMsFormsFileOutputPath(IntPtr window, string title, string value)
    {
        if (String.IsNullOrWhiteSpace(value)) throw new InvalidOperationException("File output path is required");
        SetMsFormsValue(window, 0, value);
    }

    private static bool ClickMsFormsFileOutputPreview(ProbeState state, IntPtr window, string title)
    {
        return QueueMsFormsDefaultAction(state, window, "미리보기");
    }

    private static bool ClickMsFormsFileOutputSubmit(ProbeState state, IntPtr window, string title)
    {
        if (String.Equals(title, "내엑셀 - PDF", StringComparison.Ordinal))
            return QueueMsFormsDefaultAction(state, window, "실행");
        return QueueMsFormsDefaultAction(state, window, "저장");
    }

    private static void HandleWindow(ProbeState state, IntPtr window)
    {
        var windowClass = ClassName(window);
        if (String.Equals(windowClass, "XLMAIN", StringComparison.OrdinalIgnoreCase)) return;
        // Office ribbon tooltips are transient, have no route identity, and may
        // release an HWND that MSForms immediately reuses. Tracking/closing
        // them can therefore hide a real popup under the old Seen key.
        if (String.Equals(windowClass, "Net UI Tool Window", StringComparison.OrdinalIgnoreCase)) return;
        var key = "hwnd:" + window.ToInt64().ToString();
        var firstSeen = false;
        lock (state.Gate)
        {
            firstSeen = state.Seen.Add(key);
            if (firstSeen) state.Windows.Add(ObserveWindow(window));
        }
        if (!firstSeen)
        {
            if (String.Equals(state.Mode, "popup", StringComparison.Ordinal))
            {
                PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                return;
            }
            if (String.Equals(state.Mode, "approve", StringComparison.Ordinal) &&
                windowClass.StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase))
            {
                var repeatedTitle = WindowText(window);
                if (String.IsNullOrWhiteSpace(repeatedTitle)) return;
                if (String.Equals(repeatedTitle, "내엑셀 - 실행결과", StringComparison.Ordinal))
                {
                    if (!state.PlanExecuteRequested) return;
                    ClickResultClose(state, window);
                }
                else if (!String.Equals(repeatedTitle, "내엑셀 - 실행계획", StringComparison.Ordinal))
                    PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
            if (String.Equals(state.Mode, "file-output", StringComparison.Ordinal) &&
                windowClass.StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase))
            {
                var repeatedTitle = WindowText(window);
                if (String.IsNullOrWhiteSpace(repeatedTitle)) return;
                if (String.Equals(repeatedTitle, "내엑셀 - 실행결과", StringComparison.Ordinal))
                {
                    CloseFileOutputResultWhenReady(state, window);
                    return;
                }
                if (IsMsFormsFileOutput(repeatedTitle))
                {
                    if (state.FileOutputResultClosed) return;
                    if (!state.FileOutputInputRequested)
                    {
                        SetMsFormsFileOutputPath(window, repeatedTitle, state.OutputPath);
                        state.FileOutputInputRequested = true;
                        Thread.Sleep(150);
                    }
                    if (!state.FileOutputPreviewRequested)
                    {
                        if (ClickMsFormsFileOutputPreview(state, window, repeatedTitle))
                            state.FileOutputPreviewRequested = true;
                        return;
                    }
                    if (!state.FileOutputSubmitRequested)
                    {
                        if (ClickMsFormsFileOutputSubmit(state, window, repeatedTitle))
                            state.FileOutputSubmitRequested = true;
                    }
                    return;
                }
                if (!String.Equals(repeatedTitle, "내엑셀 - 실행계획", StringComparison.Ordinal))
                {
                    var repeatedChildren = ChildWindows(window);
                    if (state.FileOutputFailureObserved)
                    {
                        ClickFirstEnabled(repeatedChildren, new[] { "취소", "닫기", "Cancel", "Close" });
                        return;
                    }
                    if (!state.FileOutputPreviewRequested &&
                        ClickFirstEnabled(repeatedChildren, new[] { "미리보기" }))
                    {
                        state.FileOutputPreviewRequested = true;
                        return;
                    }
                    if (state.FileOutputPreviewRequested && !state.FileOutputSubmitRequested &&
                        ClickFirstEnabled(repeatedChildren, new[] { "저장", "실행", "Save" }))
                    {
                        state.FileOutputSubmitRequested = true;
                    }
                }
            }
            return;
        }

        SetForegroundWindow(window);
        Thread.Sleep(100);
        var children = ChildWindows(window);
        var editValue = state.InputText;
        if (String.Equals(state.Mode, "file-output", StringComparison.Ordinal) && !String.IsNullOrWhiteSpace(state.OutputPath))
            editValue = state.OutputPath;
        SetFirstEdit(children, editValue);

        if (String.Equals(state.Mode, "approve", StringComparison.Ordinal))
        {
            if (ClickFirst(children, new[] { "예", "Yes", "확인", "OK", "실행", "계획대로 실행" })) return;
            if (String.Equals(WindowText(window), "내엑셀 - 실행계획", StringComparison.Ordinal))
            {
                state.PlanExecuteRequested = true;
                ClickPlanExecute(state, window);
                return;
            }
            if (String.Equals(WindowText(window), "내엑셀 - 실행결과", StringComparison.Ordinal))
            {
                if (!state.PlanExecuteRequested) return;
                ClickResultClose(state, window);
                return;
            }
            if (IsOfficeInputDialog(window))
            {
                SetFocusedText(window, editValue);
                PostMessage(window, WM_COMMAND, new IntPtr(IDOK), IntPtr.Zero);
                PostMessage(window, WM_KEYDOWN, new IntPtr(VK_RETURN), IntPtr.Zero);
                PostMessage(window, WM_KEYUP, new IntPtr(VK_RETURN), IntPtr.Zero);
                return;
            }
            SendFocusedKey(window, VK_RETURN);
            return;
        }

        if (String.Equals(state.Mode, "popup", StringComparison.Ordinal))
        {
            // Window-open cases only need proof that the owned popup appeared.
            // Queue the form's close/cancel action without synchronously
            // entering its event handler; fall back to closing the top-level
            // form when no conventional button is exposed.
            var popupLabels = new[] { "닫기", "취소", "Close", "Cancel", "확인", "OK" };
            if (ClickFirst(children, popupLabels)) return;
            if (TryInvokePopupButton(state.ProcessId, WindowText(window), popupLabels)) return;
            PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            return;
        }

        if (String.Equals(state.Mode, "file-output", StringComparison.Ordinal))
        {
            var fileOutputTitle = WindowText(window);
            if (windowClass.StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase) &&
                String.IsNullOrWhiteSpace(fileOutputTitle)) return;
            if (String.Equals(fileOutputTitle, "내엑셀 - 실행계획", StringComparison.Ordinal))
            {
                state.PlanExecuteRequested = true;
                ClickPlanExecute(state, window);
                return;
            }
            if (String.Equals(fileOutputTitle, "내엑셀 - 실행결과", StringComparison.Ordinal))
            {
                if (!state.PlanExecuteRequested) return;
                CloseFileOutputResultWhenReady(state, window);
                return;
            }
            if (windowClass.StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase) &&
                IsMsFormsFileOutput(fileOutputTitle))
            {
                SetMsFormsFileOutputPath(window, fileOutputTitle, editValue);
                state.FileOutputInputRequested = true;
                Thread.Sleep(150);
                if (!ClickMsFormsFileOutputPreview(state, window, fileOutputTitle))
                    throw new InvalidOperationException("File output preview is unavailable");
                state.FileOutputPreviewRequested = true;
                return;
            }
            if (windowClass.StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase) &&
                ClickFirstEnabled(children, new[] { "미리보기" }))
            {
                state.FileOutputPreviewRequested = true;
                return;
            }
            if (IsOfficeInputDialog(window))
            {
                SetFocusedText(window, editValue);
                if (ClickFirstEnabled(children, new[] { "저장", "Save", "확인", "OK" }))
                    state.FileOutputSubmitRequested = true;
                else SendFocusedKey(window, VK_RETURN);
                return;
            }
            if (ClickFirstEnabled(children, new[] { "저장", "Save", "계획대로 실행", "실행", "확인", "OK", "예", "Yes" }))
            {
                if (String.Equals(fileOutputTitle, "내엑셀", StringComparison.Ordinal)) return;
                state.FileOutputSubmitRequested = true;
                if (!windowClass.StartsWith("ThunderDFrame", StringComparison.OrdinalIgnoreCase) &&
                    fileOutputTitle.StartsWith("내엑셀 - ", StringComparison.Ordinal))
                    state.FileOutputFailureObserved = true;
                return;
            }
            PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            return;
        }

        if (ClickFirst(children, new[] { "예", "Yes", "확인", "OK", "실행", "계획대로 실행" })) return;
        if (!ClickFirst(children, new[] { "닫기", "Close", "취소", "Cancel" }))
            PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }

    private static NxObservedWindow ObserveWindow(IntPtr window)
    {
        var texts = new List<string>();
        var classes = new List<string>();
        foreach (var child in ChildWindows(window))
        {
            AddDistinct(texts, WindowText(child));
            AddDistinct(classes, ClassName(child));
        }
        CaptureOwnedPopupSurface(window, texts, classes);
        return new NxObservedWindow
        {
            Handle = window.ToInt64(),
            Title = WindowText(window),
            ClassName = ClassName(window),
            ChildTexts = texts.ToArray(),
            ChildClasses = classes.ToArray()
        };
    }

    private static void CaptureOwnedPopupSurface(IntPtr window, List<string> texts, List<string> classes)
    {
        try
        {
            var root = AutomationElement.FromHandle(window);
            if (root == null) return;
            var walker = TreeWalker.RawViewWalker;
            var pending = new Queue<AutomationElement>();
            var child = walker.GetFirstChild(root);
            while (child != null && pending.Count < 256)
            {
                pending.Enqueue(child);
                child = walker.GetNextSibling(child);
            }
            var observed = 0;
            while (pending.Count > 0 && observed < 256)
            {
                var element = pending.Dequeue();
                observed++;
                AddDistinct(texts, element.Current.Name);
                AddDistinct(classes, element.Current.ClassName);
                try
                {
                    object pattern;
                    if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                        AddDistinct(texts, ((ValuePattern)pattern).Current.Value);
                }
                catch { }
                var nested = walker.GetFirstChild(element);
                while (nested != null && pending.Count + observed < 256)
                {
                    pending.Enqueue(nested);
                    nested = walker.GetNextSibling(nested);
                }
            }
        }
        catch { }
    }

    private static void ObserveNavigatorTaskPane(ProbeState state)
    {
        try
        {
            var root = AutomationElement.FromHandle(state.MainWindow);
            if (root == null) return;
            var matches = root.FindAll(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "내엑셀 Navigator"));
            foreach (AutomationElement element in matches)
            {
                var handle = new IntPtr(element.Current.NativeWindowHandle);
                var key = "navigator:" + handle.ToInt64().ToString() + ":" + element.Current.ClassName;
                lock (state.Gate)
                {
                    if (!state.Seen.Add(key)) continue;
                    var texts = new List<string>();
                    var classes = new List<string>();
                    CaptureAutomationTree(element, texts, classes);
                    state.Windows.Add(new NxObservedWindow
                    {
                        Handle = handle.ToInt64(),
                        Title = element.Current.Name ?? String.Empty,
                        ClassName = element.Current.ClassName ?? String.Empty,
                        ChildTexts = texts.ToArray(),
                        ChildClasses = classes.ToArray()
                    });
                }
            }
        }
        catch { }
    }

    private static void CaptureAutomationTree(AutomationElement root, List<string> texts, List<string> classes)
    {
        if (root == null) return;
        try
        {
            var descendants = root.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            foreach (AutomationElement element in descendants)
            {
                AddDistinct(texts, element.Current.Name);
                AddDistinct(classes, element.Current.ClassName);
                try
                {
                    object pattern;
                    if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                        AddDistinct(texts, ((ValuePattern)pattern).Current.Value);
                    if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern) &&
                        ((SelectionItemPattern)pattern).Current.IsSelected)
                        AddDistinct(texts, element.Current.Name);
                }
                catch { }
            }
        }
        catch { }
    }

    private static void AddDistinct(List<string> values, string value)
    {
        if (String.IsNullOrWhiteSpace(value)) return;
        value = value.Trim();
        if (!values.Contains(value)) values.Add(value);
    }

    private static void SetFirstEdit(List<IntPtr> children, string value)
    {
        if (String.IsNullOrEmpty(value)) return;
        foreach (var child in children)
        {
            var className = ClassName(child);
            if (className.IndexOf("Edit", StringComparison.OrdinalIgnoreCase) < 0) continue;
            SendMessage(child, WM_SETTEXT, IntPtr.Zero, value);
            return;
        }
    }

    private static bool ClickFirst(List<IntPtr> children, string[] labels)
    {
        foreach (var wanted in labels)
        {
            foreach (var child in children)
            {
                var className = ClassName(child);
                if (className.IndexOf("Button", StringComparison.OrdinalIgnoreCase) < 0) continue;
                var caption = NormalizeButtonText(WindowText(child));
                if (!String.Equals(caption, wanted, StringComparison.OrdinalIgnoreCase)) continue;
                PostMessage(child, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return true;
            }
        }
        return false;
    }

    private static bool ClickFirstEnabled(List<IntPtr> children, string[] labels)
    {
        foreach (var wanted in labels)
        {
            foreach (var child in children)
            {
                var className = ClassName(child);
                if (className.IndexOf("Button", StringComparison.OrdinalIgnoreCase) < 0 || !IsWindowEnabled(child)) continue;
                var caption = NormalizeButtonText(WindowText(child));
                if (!String.Equals(caption, wanted, StringComparison.OrdinalIgnoreCase)) continue;
                PostMessage(child, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return true;
            }
        }
        return false;
    }

    private static string NormalizeButtonText(string text)
    {
        text = (text ?? String.Empty).Replace("&", String.Empty).Trim();
        var suffixStart = text.LastIndexOf('(');
        if (suffixStart > 0 && text.EndsWith(")", StringComparison.Ordinal))
        {
            var accelerator = text.Substring(suffixStart + 1, text.Length - suffixStart - 2);
            if (accelerator.Length <= 2) text = text.Substring(0, suffixStart).Trim();
        }
        return text;
    }

    private static void SendFocusedKey(IntPtr window, int virtualKey)
    {
        uint processId;
        var threadId = GetWindowThreadProcessId(window, out processId);
        var information = new GuiThreadInfo { Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo)) };
        var target = window;
        if (threadId != 0 && GetGUIThreadInfo(threadId, ref information) && information.Focus != IntPtr.Zero)
            target = information.Focus;
        PostMessage(target, WM_KEYDOWN, new IntPtr(virtualKey), IntPtr.Zero);
        PostMessage(target, WM_KEYUP, new IntPtr(virtualKey), IntPtr.Zero);
    }

    private static void SetFocusedText(IntPtr window, string value)
    {
        uint processId;
        var threadId = GetWindowThreadProcessId(window, out processId);
        var information = new GuiThreadInfo { Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo)) };
        if (threadId == 0 || !GetGUIThreadInfo(threadId, ref information) || information.Focus == IntPtr.Zero) return;
        SendMessage(information.Focus, WM_SETTEXT, IntPtr.Zero, value ?? String.Empty);
    }

    private static bool IsOfficeInputDialog(IntPtr window)
    {
        return String.Equals(WindowText(window), "LHexcel", StringComparison.Ordinal) ||
            ClassName(window).StartsWith("bosa_sdm", StringComparison.OrdinalIgnoreCase);
    }

    private static void CloseFileOutputResultWhenReady(ProbeState state, IntPtr window)
    {
        if (!state.PlanExecuteRequested || state.ResultCloseRequested) return;
        state.ResultCloseRequested = true;
        try
        {
            var started = Environment.TickCount;
            while (!state.StopRequested && IsWindowVisible(window) &&
                String.Equals(WindowText(window), "내엑셀 - 실행결과", StringComparison.Ordinal) &&
                unchecked(Environment.TickCount - started) < 10000)
            {
                if (String.IsNullOrWhiteSpace(state.OutputPath) || File.Exists(state.OutputPath) || Directory.Exists(state.OutputPath))
                {
                    Thread.Sleep(500);
                    break;
                }
                Thread.Sleep(100);
            }
            if (state.StopRequested || !IsWindowVisible(window) ||
                !String.Equals(WindowText(window), "내엑셀 - 실행결과", StringComparison.Ordinal))
            {
                state.ResultCloseRequested = false;
                return;
            }
            ClickResultClose(state, window);
            var closeDeadline = unchecked(Environment.TickCount + 2000);
            while (IsWindowVisible(window) && String.Equals(WindowText(window), "내엑셀 - 실행결과", StringComparison.Ordinal) &&
                unchecked(closeDeadline - Environment.TickCount) > 0) Thread.Sleep(50);
            if (IsWindowVisible(window) && String.Equals(WindowText(window), "내엑셀 - 실행결과", StringComparison.Ordinal))
                state.ResultCloseRequested = false;
            else
                state.FileOutputResultClosed = true;
        }
        catch { state.ResultCloseRequested = false; }
    }

    private static void ClickPlanExecute(ProbeState state, IntPtr window)
    {
        if (!QueueMsFormsDefaultAction(state, window, "계획대로 실행"))
            throw new InvalidOperationException("Plan execute is unavailable");
    }

    private static void ClickResultClose(ProbeState state, IntPtr window)
    {
        if (!PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero))
            throw new InvalidOperationException("Result close is unavailable");
    }

    private static bool TryInvokePopupButton(int processId, string windowTitle, string[] labels)
    {
        foreach (var label in labels)
            if (TryInvokeNamedButton(processId, windowTitle, label)) return true;
        return false;
    }

    private static bool TryInvokeNamedButton(int processId, string windowTitle, string label)
    {
        try
        {
            var containerCondition = new AndCondition(
                new PropertyCondition(AutomationElement.ProcessIdProperty, processId),
                new PropertyCondition(AutomationElement.NameProperty, windowTitle),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
            var buttonCondition = new AndCondition(
                new PropertyCondition(AutomationElement.ProcessIdProperty, processId),
                new PropertyCondition(AutomationElement.NameProperty, label),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            AutomationElement button = null;
            EnumWindows((candidateWindow, parameter) =>
            {
                uint candidateProcessId;
                GetWindowThreadProcessId(candidateWindow, out candidateProcessId);
                if (candidateProcessId != (uint)processId ||
                    !String.Equals(ClassName(candidateWindow), "XLMAIN", StringComparison.OrdinalIgnoreCase)) return true;
                try
                {
                    var root = AutomationElement.FromHandle(candidateWindow);
                    if (root != null)
                    {
                        var container = root.FindFirst(TreeScope.Descendants, containerCondition);
                        if (container != null) button = container.FindFirst(TreeScope.Descendants, buttonCondition);
                    }
                }
                catch { }
                return button == null;
            }, IntPtr.Zero);
            if (button == null) return false;
            object pattern;
            if (!button.TryGetCurrentPattern(InvokePattern.Pattern, out pattern)) return false;
            var invoke = (InvokePattern)pattern;
            ThreadPool.QueueUserWorkItem(delegate
            {
                try { invoke.Invoke(); }
                catch { }
            });
            return true;
        }
        catch { return false; }
    }

    private static List<IntPtr> ChildWindows(IntPtr parent)
    {
        var result = new List<IntPtr>();
        EnumChildWindows(parent, (window, parameter) => { result.Add(window); return true; }, IntPtr.Zero);
        return result;
    }

    private static string WindowText(IntPtr window)
    {
        var value = new StringBuilder(2048);
        GetWindowText(window, value, value.Capacity);
        return value.ToString();
    }

    private static string ClassName(IntPtr window)
    {
        var value = new StringBuilder(256);
        GetClassName(window, value, value.Capacity);
        return value.ToString();
    }
}
