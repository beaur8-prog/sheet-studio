using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

// Query the same association handlers used by Explorer's Open With menu.
// This helper runs only on the disposable installer CI runner.
public static class SheetFileSupport {
    [ComImport, Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellLinkW {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder path, int capacity, IntPtr findData, uint flags);
    }

    public static string ShortcutTarget(string filename) {
        var link = Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("00021401-0000-0000-C000-000000000046")));
        try {
            ((IPersistFile)link).Load(filename, 0);
            var path = new StringBuilder(32768);
            ((IShellLinkW)link).GetPath(path, path.Capacity, IntPtr.Zero, 4);
            return path.ToString();
        } finally { Marshal.ReleaseComObject(link); }
    }

    [ComImport, Guid("F04061AC-1659-4A3F-A954-775AA57FC083"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAssocHandler {
        void GetName([MarshalAs(UnmanagedType.LPWStr)] out string name);
        void GetUIName([MarshalAs(UnmanagedType.LPWStr)] out string name);
    }
    [ComImport, Guid("973810AE-9599-4B88-9E4D-6EE98C9552DA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IEnumAssocHandlers {
        [PreserveSig] int Next(uint count, [MarshalAs(UnmanagedType.Interface)] out IAssocHandler handler, out uint fetched);
    }
    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    static extern int SHAssocEnumHandlers(string extension, uint filter, out IEnumAssocHandlers handlers);

    public static string[] Applications(string extension) {
        IEnumAssocHandlers handlers;
        Marshal.ThrowExceptionForHR(SHAssocEnumHandlers(extension, 0, out handlers));
        var names = new List<string>();
        if (handlers == null) return names.ToArray();
        try {
            IAssocHandler handler; uint fetched;
            while (handlers.Next(1, out handler, out fetched) == 0 && fetched == 1) {
                try { string name; handler.GetUIName(out name); names.Add(name); }
                finally { Marshal.ReleaseComObject(handler); }
            }
        } finally { Marshal.ReleaseComObject(handlers); }
        return names.ToArray();
    }
}
