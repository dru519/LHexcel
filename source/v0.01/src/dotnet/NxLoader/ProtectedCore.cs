using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace LH.NxHost {
    [ComVisible(false)]
    internal static class ProtectedCore {
        private static readonly object Gate = new object();
        private static Assembly loaded;

        internal static object Create(string name) {
            lock (Gate) {
                string directory = Path.GetDirectoryName(typeof(ProtectedCore).Assembly.Location);
                string path = Path.GetFullPath(Path.Combine(directory, CorePin.FileName));
                if (!String.Equals(Path.GetDirectoryName(path), directory, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("내엑셀 핵심 모듈 경로를 확인할 수 없습니다.");
                // Keep the exact file write/delete locked across hash and load.
                using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    string actual;
                    using (var sha = SHA256.Create())
                        actual = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "").ToLowerInvariant();
                    if (!String.Equals(actual, CorePin.Sha256, StringComparison.Ordinal))
                        throw new InvalidOperationException("내엑셀 핵심 모듈이 변경되었습니다. 원본 배포본을 다시 설치해 주세요.");
                    if (loaded == null) {
                        Assembly candidate = Assembly.LoadFrom(path);
                        // LoadFrom may bind an already loaded assembly of the same
                        // identity. Check both its location and compiled module ID.
                        if (!String.Equals(candidate.Location, path, StringComparison.OrdinalIgnoreCase) ||
                            candidate.ManifestModule.ModuleVersionId.ToString("D") != CorePin.ModuleId)
                            throw new InvalidOperationException("내엑셀 핵심 모듈 버전이 일치하지 않습니다. Excel을 다시 시작해 주세요.");
                        loaded = candidate;
                    } else if (loaded.ManifestModule.ModuleVersionId.ToString("D") != CorePin.ModuleId) {
                        throw new InvalidOperationException("내엑셀 핵심 모듈 버전이 일치하지 않습니다.");
                    }
                    switch (name) {
                        case "BridgeService": case "HwpxExportService": case "PicturePreviewService":
                        case "WorkbookCompareService": case "WorkbookCompareResultsService":
                        case "NxHostAddIn": case "NavigatorPane": break;
                        default: throw new InvalidOperationException("지원하지 않는 내엑셀 구성요소입니다.");
                    }
                    return Activator.CreateInstance(loaded.GetType("LH.NxHost." + name, true, false));
                }
            }
        }

        internal static object Call(object target, string name, params object[] args) {
            try {
                return target.GetType().InvokeMember(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod,
                    null, target, args);
            } catch (TargetInvocationException error) {
                if (error.InnerException != null) ExceptionDispatchInfo.Capture(error.InnerException).Throw();
                throw;
            }
        }
        internal static object Get(object target, string name) {
            return target.GetType().GetProperty(name).GetValue(target, null);
        }
    }
}
