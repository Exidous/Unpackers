using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace NonIntrusive
{
    public class Win32
    {
        #region DLL Imports
        [DllImport("kernel32.dll")]
        public static extern int CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll")]
        public static extern bool Module32First(IntPtr hSnapshot, ref MODULEENTRY32 lpme);
        [DllImport("kernel32.dll")]
        public static extern bool Module32Next(IntPtr hSnapshot, ref MODULEENTRY32 lpme);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr CreateToolhelp32Snapshot(SnapshotFlags dwFlags, uint th32ProcessID);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        public static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);
        [DllImport("kernel32.dll")]
        public static extern bool VirtualProtectEx(IntPtr hProcess, IntPtr lpAddress, UIntPtr dwSize, uint flNewProtect, out uint lpflOldProtect);
        [DllImport("kernel32.dll")]
        public static extern int VirtualQueryEx(IntPtr hProcess, IntPtr lpAddress, out MEMORY_BASIC_INFORMATION lpBuffer, uint dwLength);
        [DllImport("kernel32.dll")]
        public static extern int CreateProcess(string lpApplicationName, string lpCommandLine, [MarshalAs(UnmanagedType.Struct)] ref  SECURITY_ATTRIBUTES lpProcessAttributes, [MarshalAs(UnmanagedType.Struct)] ref  SECURITY_ATTRIBUTES lpThreadAttributes, int bInheritHandles, int dwCreationFlags, int lpEnvironment, string lpCurrentDriectory, [MarshalAs(UnmanagedType.Struct)] ref  STARTUPINFO lpStartupInfo, [MarshalAs(UnmanagedType.Struct)] ref  PROCESS_INFORMATION lpProcessInformation);
        [DllImport("kernel32.dll")]
        public static extern int ResumeThread(IntPtr hThread);
        [DllImport("kernel32.dll")]
        public static extern int SuspendThread(IntPtr hThread);
        [DllImport("kernel32.dll")]
        public static extern bool GetThreadContext(IntPtr hThread, ref CONTEXT lpContext);
        [DllImport("kernel32.dll")]
        public static extern int SetThreadContext(IntPtr hThread, [MarshalAs(UnmanagedType.Struct)] ref  CONTEXT lpContext);
        [DllImport("kernel32.dll")]
        public static extern int ReadProcessMemory(int hProcess, int lpBaseAddress, byte[] lpBuffer, int nSize, ref int lpNumberOfbytesWritten);
        [DllImport("kernel32.dll")]
        public static extern int WriteProcessMemory(int hProcess, int lpBaseAddress, byte[] lpBuffer, int nSize, ref int lpNumberOfbytesWritten);
        [DllImport("kernel32.dll")]
        public static extern void Sleep(int dwMilliseconds);
        [DllImport("kernel32.dll")]
        public static extern int GetTickCount();
        [DllImport("kernel32.dll")]
        public static extern int GetProcAddress(int hModule, string lpProcName);
        [DllImport("kernel32.dll")]
        public static extern int GetModuleHandle(string lpModuleName);
        [DllImport("kernel32.dll")]
        public static extern int VirtualQueryEx(int hProcess, int lpAddress, [MarshalAs(UnmanagedType.Struct)] ref  MEMORY_BASIC_INFORMATION lpBuffer, int dwLength);
        [DllImport("kernel32.dll")]
        public static extern int TerminateProcess(int hProcess, int uExitCode);
        [DllImport("kernel32.dll")]
        public static extern IntPtr OpenThread(int dwDesiredAccess, bool bInheritHandle, uint dwThreadId);
        #endregion
        #region Structs/Enums
        public struct MODULEENTRY32
        {
            const int MAX_PATH = 260;
            public UInt32 dwSize;
            public UInt32 th32ModuleID;
            public UInt32 th32ProcessID;
            public IntPtr GlblcntUsage;
            public UInt32 ProccntUsage;
            public IntPtr modBaseAddr;
            public UInt32 modBaseSize;
            public IntPtr hModule;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
            public string szModule;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
            public string szExePath;
            public UInt32 dwFlags;
        }
        [Flags]
        public enum SnapshotFlags : uint
        {
            HeapList = 0x00000001,
            Process = 0x00000002,
            Thread = 0x00000004,
            Module = 0x00000008,
            Module32 = 0x00000010,
            Inherit = 0x80000000,
            All = 0x0000001F,
            NoHeaps = 0x40000000
        }
        public enum AllocationProtectEnum : uint
        {
            PAGE_EXECUTE = 0x00000010,
            PAGE_EXECUTE_READ = 0x00000020,
            PAGE_EXECUTE_READWRITE = 0x00000040,
            PAGE_EXECUTE_WRITECOPY = 0x00000080,
            PAGE_NOACCESS = 0x00000001,
            PAGE_READONLY = 0x00000002,
            PAGE_READWRITE = 0x00000004,
            PAGE_WRITECOPY = 0x00000008,
            PAGE_GUARD = 0x00000100,
            PAGE_NOCACHE = 0x00000200,
            PAGE_WRITECOMBINE = 0x00000400
        }

        public enum StateEnum : uint
        {
            MEM_COMMIT = 0x1000,
            MEM_FREE = 0x10000,
            MEM_RESERVE = 0x2000
        }

        public enum TypeEnum : uint
        {
            MEM_IMAGE = 0x1000000,
            MEM_MAPPED = 0x40000,
            MEM_PRIVATE = 0x20000
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MEMORY_BASIC_INFORMATION
        {
            public IntPtr BaseAddress;
            public IntPtr AllocationBase;
            public AllocationProtectEnum AllocationProtect;
            public uint RegionSize;
            public StateEnum State;
            public AllocationProtectEnum Protect;
            public TypeEnum Type;
        }
        public enum CONTEXT_FLAGS : uint
        {
            CONTEXT_i386 = 0x10000,
            CONTEXT_i486 = 0x10000,   //  same as i386
            CONTEXT_CONTROL = CONTEXT_i386 | 0x01, // SS:SP, CS:IP, FLAGS, BP
            CONTEXT_INTEGER = CONTEXT_i386 | 0x02, // AX, BX, CX, DX, SI, DI
            CONTEXT_SEGMENTS = CONTEXT_i386 | 0x04, // DS, ES, FS, GS
            CONTEXT_FLOATING_POINT = CONTEXT_i386 | 0x08, // 387 state
            CONTEXT_DEBUG_REGISTERS = CONTEXT_i386 | 0x10, // DB 0-3,6,7
            CONTEXT_EXTENDED_REGISTERS = CONTEXT_i386 | 0x20, // cpu specific extensions
            CONTEXT_FULL = CONTEXT_CONTROL | CONTEXT_INTEGER | CONTEXT_SEGMENTS,
            CONTEXT_ALL = CONTEXT_CONTROL | CONTEXT_INTEGER | CONTEXT_SEGMENTS | CONTEXT_FLOATING_POINT | CONTEXT_DEBUG_REGISTERS | CONTEXT_EXTENDED_REGISTERS
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SECURITY_ATTRIBUTES
        {
            public int nLength;
            public int lpSecurityDescriptor;
            public int bInheritHandle;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct STARTUPINFO
        {
            public int cb;
            public string lpReserved;
            public string lpDesktop;
            public string lpTitle;
            public int dwX;
            public int dwY;
            public int dwXSize;
            public int dwYSize;
            public int dwXCountChars;
            public int dwYCountChars;
            public int dwFillAttribute;
            public int dwFlags;
            public int wShowWindow;
            public int cbReserved2;
            public int lpReserved2;
            public int hStdInput;
            public int hStdOutput;
            public int hStdError;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public int hProcess;
            public int hThread;
            public int dwProcessId;
            public int dwThreadId;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct FLOATING_SAVE_AREA
        {
            public uint ControlWord;
            public uint StatusWord;
            public uint TagWord;
            public uint ErrorOffset;
            public uint ErrorSelector;
            public uint DataOffset;
            public uint DataSelector;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 80)]
            public byte[] RegisterArea;
            public uint Cr0NpxState;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct CONTEXT
        {
            public uint ContextFlags; //set this to an appropriate value 
            // Retrieved by CONTEXT_DEBUG_REGISTERS
            public uint Dr0;
            public uint Dr1;
            public uint Dr2;
            public uint Dr3;
            public uint Dr6;
            public uint Dr7;
            // Retrieved by CONTEXT_FLOATING_POINT
            public FLOATING_SAVE_AREA FloatSave;
            // Retrieved by CONTEXT_SEGMENTS
            public uint SegGs;
            public uint SegFs;
            public uint SegEs;
            public uint SegDs;
            // Retrieved by CONTEXT_INTEGER
            public uint Edi;
            public uint Esi;
            public uint Ebx;
            public uint Edx;
            public uint Ecx;
            public uint Eax;
            // Retrieved by CONTEXT_CONTROL
            public uint Ebp;
            public uint Eip;
            public uint SegCs;
            public uint EFlags;
            public uint Esp;
            public uint SegSs;
            // Retrieved by CONTEXT_EXTENDED_REGISTERS
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 512)]
            public byte[] ExtendedRegisters;

        }
        #endregion

        public const int CREATE_SUSPENDED = 0x4;
        public const int GET_CONTEXT = 0x8;
        public const int SET_CONTEXT = 0x0010;

        
    }
}
