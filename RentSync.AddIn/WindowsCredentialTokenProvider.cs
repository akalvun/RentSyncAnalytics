using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using RentSync.Core.Services;

namespace RentSync.AddIn
{
    /// <summary>
    /// Reads the API token from Windows Credential Manager (CredRead Win32
    /// API). Nothing sensitive ever lands in config files or the registry;
    /// the credential is stored per-user, DPAPI-protected by the OS.
    ///
    /// To store the token, users run once in cmd:
    ///   cmdkey /generic:RentSync.ApiToken /user:api /pass:THEIR_TOKEN
    /// or use the Settings dialog, which calls CredWrite.
    /// </summary>
    public sealed class WindowsCredentialTokenProvider : IAuthTokenProvider
    {
        public const string TargetName = "RentSync.ApiToken";

        public Task<string> GetTokenAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult(ReadCredential(TargetName));

        public static string ReadCredential(string target)
        {
            if (!CredRead(target, CredType.Generic, 0, out var handle))
                return null; // no credential stored -> anonymous demo mode

            try
            {
                var cred = Marshal.PtrToStructure<NativeCredential>(handle);
                if (cred.CredentialBlob == IntPtr.Zero || cred.CredentialBlobSize == 0)
                    return null;
                var bytes = new byte[cred.CredentialBlobSize];
                Marshal.Copy(cred.CredentialBlob, bytes, 0, bytes.Length);
                return Encoding.Unicode.GetString(bytes);
            }
            finally
            {
                CredFree(handle);
            }
        }

        public static void WriteCredential(string target, string secret)
        {
            var bytes = Encoding.Unicode.GetBytes(secret);
            var blob = Marshal.AllocHGlobal(bytes.Length);
            try
            {
                Marshal.Copy(bytes, 0, blob, bytes.Length);
                var cred = new NativeCredential
                {
                    Type = CredType.Generic,
                    TargetName = target,
                    UserName = "api",
                    CredentialBlob = blob,
                    CredentialBlobSize = (uint)bytes.Length,
                    Persist = 2 // CRED_PERSIST_LOCAL_MACHINE
                };
                if (!CredWrite(ref cred, 0))
                    throw new InvalidOperationException(
                        "CredWrite failed: " + Marshal.GetLastWin32Error());
            }
            finally
            {
                Marshal.FreeHGlobal(blob);
            }
        }

        // ---------------- Win32 ----------------

        private enum CredType : uint { Generic = 1 }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NativeCredential
        {
            public uint Flags;
            public CredType Type;
            public string TargetName;
            public string Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public uint CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint Persist;
            public uint AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CredReadW")]
        private static extern bool CredRead(string target, CredType type, int flags, out IntPtr credential);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CredWriteW")]
        private static extern bool CredWrite(ref NativeCredential credential, uint flags);

        [DllImport("advapi32.dll")]
        private static extern void CredFree(IntPtr buffer);
    }
}
