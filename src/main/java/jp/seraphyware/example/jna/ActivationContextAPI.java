package jp.seraphyware.example.jna;

import java.util.Arrays;
import java.util.List;
import java.util.concurrent.Callable;

import com.sun.jna.Native;
import com.sun.jna.Structure;
import com.sun.jna.WString;
import com.sun.jna.platform.win32.BaseTSD.ULONG_PTR;
import com.sun.jna.platform.win32.BaseTSD.ULONG_PTRByReference;
import com.sun.jna.platform.win32.Kernel32;
import com.sun.jna.platform.win32.Win32Exception;
import com.sun.jna.platform.win32.WinDef.DWORD;
import com.sun.jna.platform.win32.WinDef.HMODULE;
import com.sun.jna.platform.win32.WinDef.ULONG;
import com.sun.jna.platform.win32.WinDef.USHORT;
import com.sun.jna.platform.win32.WinNT;
import com.sun.jna.platform.win32.WinNT.HANDLE;
import com.sun.jna.win32.StdCallLibrary;
import com.sun.jna.win32.W32APIOptions;

/**
 * JNAのKernel32では定義されていない、Activation ContextまわりのAPIの追加分.
 *
 * https://msdn.microsoft.com/en-us/library/windows/desktop/aa374153.aspx
 */
public interface ActivationContextAPI extends StdCallLibrary {

	ActivationContextAPI INSTANCE = (ActivationContextAPI) Native.loadLibrary(
			"Kernel32", ActivationContextAPI.class, W32APIOptions.UNICODE_OPTIONS);

	/**
	 * ACTCTX構造体
	 */
	public static class ACTCTX extends Structure {

		public ULONG cbSize;

		public DWORD dwFlags = new DWORD(0);

		public WString lpSource; // マニフェストファイル(Unicode)

		public USHORT wProcessorArchitecture = new USHORT(0);

		public USHORT wLangId = new USHORT(0);

		public String lpAssemblyDirectory;

		public String lpResourceName;

		public String lpApplicationName;

		public HMODULE hModule;

		protected List<String> getFieldOrder() {
			return Arrays.asList(new String[] { "cbSize", "dwFlags", "lpSource", "wProcessorArchitecture",
					"wLangId", "lpAssemblyDirectory", "lpResourceName", "lpApplicationName", "hModule" });
		}

		public ACTCTX() {
			cbSize = new ULONG(size());
		}
	}

	/**
	 * The CreateActCtx function creates an activation context.
	 *
	 * https://msdn.microsoft.com/en-us/library/windows/desktop/aa375125.aspx
	 *
	 * @param pActCtx
	 * 	Pointer to an ACTCTX structure that contains information about the activation context to be created.
	 *
	 * @return
	 * 	If the function succeeds, it returns a handle to the returned activation context.
	 * 	Otherwise, it returns INVALID_HANDLE_VALUE.
	 */
	HANDLE CreateActCtx(ActivationContextAPI.ACTCTX pActCtx);

	/**
	 * The ReleaseActCtx function decrements the reference count of the specified activation context.
	 *
	 * https://msdn.microsoft.com/en-us/library/windows/desktop/aa375713.aspx
	 *
	 * @param handle
	 * 	Handle to the ACTCTX structure that contains information on the activation context for
	 * 	which the reference count is to be decremented.
	 */
	void ReleaseActCtx(HANDLE handle);

	/**
	 * The GetCurrentActCtx function returns the handle to the active activation context of the calling thread.
	 *
	 * https://msdn.microsoft.com/en-us/library/windows/desktop/aa375152.aspx
	 *
	 * @param lpHandle
	 * 	Pointer to the returned ACTCTX structure that contains information on the active activation context.
	 *
	 * @return
	 * 	If the function succeeds, it returns TRUE. Otherwise, it returns FALSE.
	 */
	boolean GetCurrentActCtx(WinNT.HANDLEByReference lpHandle);

	/**
	 * The ActivateActCtx function activates the specified activation context.
	 * It does this by pushing the specified activation context to the top of the activation stack.
	 * The specified activation context is thus associated with the current thread and any appropriate
	 *  side-by-side API functions.
	 *
	 *  https://msdn.microsoft.com/en-us/library/windows/desktop/aa374151.aspx
	 *
	 * @param handle
	 * 	Handle to an ACTCTX structure that contains information on the activation context that is to be made active.
	 *
	 * @param lpCookie
	 * 	Pointer to a ULONG_PTR that functions as a cookie, uniquely identifying a specific, activated activation context.
	 *
	 * @return
	 * 	If the function succeeds, it returns TRUE. Otherwise, it returns FALSE.
	 */
	boolean ActivateActCtx(HANDLE handle, ULONG_PTRByReference lpCookie);

	/**
	 * The DeactivateActCtx function deactivates the activation context corresponding to the specified cookie.
	 *
	 * https://msdn.microsoft.com/en-us/library/windows/desktop/aa375140.aspx
	 *
	 * @param dwFlags
	 * 	Flags that indicate how the deactivation.
	 *
	 * @param cookie
	 * 	The ULONG_PTR that was passed into the call to ActivateActCtx.
	 * 	This value is used as a cookie to identify a specific activated activation context.
	 *
	 * @return
	 * 	If the function succeeds, it returns TRUE. Otherwise, it returns FALSE.
	 */
	boolean DeactivateActCtx(DWORD dwFlags, ULONG_PTR cookie);

	int ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID = 1;
	int ACTCTX_FLAG_LANGID_VALID = 2;
	int ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 4;
	int ACTCTX_FLAG_RESOURCE_NAME_VALID = 8;
	int ACTCTX_FLAG_SET_PROCESS_DEFAULT = 16;
	int ACTCTX_FLAG_APPLICATION_NAME_VALID = 32;
	int ACTCTX_FLAG_HMODULE_VALID = 128;

	int PROCESSOR_ARCHITECTURE_AMD64 = 9;
	int PROCESSOR_ARCHITECTURE_ARM = 5;
	int PROCESSOR_ARCHITECTURE_ARM64 = 12;
	int PROCESSOR_ARCHITECTURE_IA64 = 6;
	int PROCESSOR_ARCHITECTURE_INTEL = 0; // x86
	int PROCESSOR_ARCHITECTURE_UNKNOWN = 0xffff;

	int DEACTIVATE_ACTCTX_FLAG_FORCE_EARLY_DEACTIVATION = 1;

	/**
	 * マニフェストファイルを指定して、アクティベーションコンテキストを作成、有効化し、
	 * 指定されたRunnableを実行したあと、コンテキストを無効化して破棄して終了する。
	 * (タスク内で発生した例外はRuntimeExceptionにラップされる。)
	 * @param manifestFile マニフェストファイルのパス
	 * @param r アクティベーションコンテキスト中で実行するタスク
	 */
	static <T> T doActivate(String manifestFile, Callable<T> r) {
		Kernel32 kernel32 = Kernel32.INSTANCE;

		ACTCTX actctx = new ACTCTX();
		actctx.lpSource = new WString(manifestFile);
		HANDLE handle = INSTANCE.CreateActCtx(actctx);

		if (handle == Kernel32.INVALID_HANDLE_VALUE) {
			throw new Win32Exception(kernel32.GetLastError());
		}
		try {
			ULONG_PTRByReference lpCookie = new ULONG_PTRByReference();
			if (!INSTANCE.ActivateActCtx(handle, lpCookie)) {
				throw new Win32Exception(kernel32.GetLastError());
			}
			ULONG_PTR cookie = lpCookie.getValue();
			try {
				return r.call();

			} catch (RuntimeException | Error ex) {
				throw ex;

			} catch (Exception ex) {
				throw new RuntimeException(ex);

			} finally {
				INSTANCE.DeactivateActCtx(new DWORD(0), cookie);
			}
		} finally {
			INSTANCE.ReleaseActCtx(handle);
		}
	}

	static void doActivate(String manifestFile, Runnable r) {
		doActivate(manifestFile, () -> {
			r.run();
			return null;
		});
	}
}
