using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Cirya.WordPdf {
    public static class WindowsNative {
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetIdsOfNames(IntPtr dispatch, ref Guid iid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] names,
            uint count, uint locale, [Out] int[] ids);
        public static bool HasMethod(object target, string method) {
            IntPtr dispatch = Marshal.GetIDispatchForObject(target);
            try {
                IntPtr table = Marshal.ReadIntPtr(dispatch);
                IntPtr address = Marshal.ReadIntPtr(table, 5 * IntPtr.Size);
                var getIds = (GetIdsOfNames)Marshal.GetDelegateForFunctionPointer(address, typeof(GetIdsOfNames));
                Guid iid = Guid.Empty;
                int result = getIds(dispatch, ref iid, new string[] { method }, 1, 0, new int[1]);
                if (result == unchecked((int)0x80020006) || result == unchecked((int)0x80020003)) return false;
                Marshal.ThrowExceptionForHR(result);
                return true;
            } finally { Marshal.Release(dispatch); }
        }
        public static object Invoke(object target, string method, object[] arguments) {
            return target.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, target, arguments);
        }
        public static void Publish(string source, string destination, bool overwrite) {
            if (overwrite && File.Exists(destination)) File.Replace(source, destination, null);
            else File.Move(source, destination);
        }
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr data);
        private delegate bool EnumCallback(IntPtr window, IntPtr data);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out int pid);
        public static int TaggedWordPid(string marker) {
            var pids = new HashSet<int>();
            if (!EnumWindows(delegate(IntPtr window, IntPtr unused) {
                var name = new StringBuilder(256);
                GetClassName(window, name, name.Capacity);
                if (name.ToString() == "OpusApp") {
                    var title = new StringBuilder(2048);
                    GetWindowText(window, title, title.Capacity);
                    if (title.ToString().Contains(marker)) {
                        int pid;
                        GetWindowThreadProcessId(window, out pid);
                        pids.Add(pid);
                    }
                }
                return true;
            }, IntPtr.Zero)) throw new InvalidOperationException("Cannot enumerate Word windows.");
            if (pids.Count != 1) throw new InvalidOperationException("Cannot uniquely identify Word's process.");
            foreach (int pid in pids) return pid;
            throw new InvalidOperationException("Word window not found.");
        }
        public static double Created(Process process) {
            return (process.StartTime.ToUniversalTime() - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
        }
        public static Process Matching(int pid, double created, string name) {
            Process process;
            try { process = Process.GetProcessById(pid); }
            catch (ArgumentException) { return null; }
            try {
                if (process.HasExited || Math.Abs(Created(process) - created) > 0.01 ||
                    !String.Equals(process.ProcessName + ".exe", name, StringComparison.OrdinalIgnoreCase)) {
                    process.Dispose();
                    return null;
                }
                return process;
            } catch { process.Dispose(); throw; }
        }
        public static void StopOwned(int pid, double created, string name) {
            if (!String.Equals(name, "WINWORD.EXE", StringComparison.OrdinalIgnoreCase) &&
                !String.Equals(name, "powershell.exe", StringComparison.OrdinalIgnoreCase) &&
                !String.Equals(name, "python.exe", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Refusing to stop an unexpected process type.");
            using (Process process = Matching(pid, created, name)) {
                if (process == null) return;
                try { process.Kill(); } catch (InvalidOperationException) { }
                if (!process.WaitForExit(5000)) throw new InvalidOperationException("Owned process did not exit.");
            }
        }
        public static void WatchParent(int pid) {
            double created;
            string name;
            using (var parent = Process.GetProcessById(pid)) {
                created = Created(parent);
                name = parent.ProcessName + ".exe";
            }
            var watcher = new Thread(delegate() {
                while (true) {
                    try {
                        using (var parent = Matching(pid, created, name)) {
                            if (parent == null) Environment.Exit(70);
                        }
                    } catch { Environment.Exit(70); }
                    Thread.Sleep(200);
                }
            });
            watcher.IsBackground = true;
            watcher.Start();
        }
    }
    // Acquisition and disposal must occur on the same OS/managed thread.
    public sealed class MutexLease : IDisposable {
        private Mutex mutex;
        private int ownerThread;
        public bool Abandoned { get; private set; }
        public MutexLease(string name, int milliseconds) {
            mutex = new Mutex(false, name);
            try {
                try {
                    if (!mutex.WaitOne(milliseconds)) throw new TimeoutException("queue_timeout: Timed out waiting for Word export lock.");
                } catch (AbandonedMutexException) { Abandoned = true; }
                ownerThread = Thread.CurrentThread.ManagedThreadId;
            } catch { mutex.Dispose(); mutex = null; throw; }
        }
        public void Dispose() {
            if (mutex == null) return;
            if (ownerThread != Thread.CurrentThread.ManagedThreadId)
                throw new InvalidOperationException("Mutex must be released on its acquiring thread.");
            mutex.ReleaseMutex();
            mutex.Dispose();
            mutex = null;
        }
    }
    public sealed class OwnedJob : IDisposable {
        [StructLayout(LayoutKind.Sequential)] private struct BasicLimits {
            public long ProcessTime, JobTime;
            public uint Flags;
            public UIntPtr MinWorkingSet, MaxWorkingSet;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass, SchedulingClass;
        }
        [StructLayout(LayoutKind.Sequential)] private struct IoCounters {
            public ulong ReadOperations, WriteOperations, OtherOperations, ReadBytes, WriteBytes, OtherBytes;
        }
        [StructLayout(LayoutKind.Sequential)] private struct ExtendedLimits {
            public BasicLimits Basic;
            public IoCounters Io;
            public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateJobObject(IntPtr attributes, string name);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetInformationJobObject(IntPtr job, int infoClass, ref ExtendedLimits limits, uint length);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        private IntPtr handle;
        public OwnedJob(int pid) {
            handle = CreateJobObject(IntPtr.Zero, null);
            if (handle == IntPtr.Zero) throw new System.ComponentModel.Win32Exception();
            try {
                var limits = new ExtendedLimits();
                limits.Basic.Flags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                if (!SetInformationJobObject(handle, 9, ref limits, (uint)Marshal.SizeOf(typeof(ExtendedLimits))))
                    throw new System.ComponentModel.Win32Exception();
                IntPtr process = OpenProcess(0x100101, false, pid); // SYNCHRONIZE | SET_QUOTA | TERMINATE
                if (process == IntPtr.Zero) throw new System.ComponentModel.Win32Exception();
                try {
                    if (!AssignProcessToJobObject(handle, process)) throw new System.ComponentModel.Win32Exception();
                } finally { CloseHandle(process); }
            } catch { Dispose(); throw; }
        }
        public void Dispose() {
            if (handle != IntPtr.Zero) { CloseHandle(handle); handle = IntPtr.Zero; }
        }
    }
}
