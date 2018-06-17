package jp.seraphyware.example.jna;

import java.util.concurrent.ConcurrentLinkedDeque;

import com.sun.jna.Pointer;
import com.sun.jna.WString;
import com.sun.jna.platform.win32.Guid.IID;
import com.sun.jna.platform.win32.Guid.REFIID;
import com.sun.jna.platform.win32.OaIdl.DISPID;
import com.sun.jna.platform.win32.OaIdl.DISPIDByReference;
import com.sun.jna.platform.win32.OaIdl.EXCEPINFO;
import com.sun.jna.platform.win32.OaIdl.VARIANT_BOOLByReference;
import com.sun.jna.platform.win32.OleAuto.DISPPARAMS;
import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.WinDef.DWORD;
import com.sun.jna.platform.win32.WinDef.DWORDByReference;
import com.sun.jna.platform.win32.WinDef.LCID;
import com.sun.jna.platform.win32.WinDef.UINT;
import com.sun.jna.platform.win32.WinDef.UINTByReference;
import com.sun.jna.platform.win32.WinDef.WORD;
import com.sun.jna.platform.win32.WinError;
import com.sun.jna.platform.win32.WinNT.HRESULT;
import com.sun.jna.platform.win32.COM.COMLateBindingObject;
import com.sun.jna.platform.win32.COM.COMUtils;
import com.sun.jna.platform.win32.COM.ConnectionPoint;
import com.sun.jna.platform.win32.COM.ConnectionPointContainer;
import com.sun.jna.platform.win32.COM.DispatchListener;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.COM.IDispatchCallback;
import com.sun.jna.platform.win32.COM.IUnknown;
import com.sun.jna.ptr.IntByReference;
import com.sun.jna.ptr.PointerByReference;

/**
 * JNAによる、MyRegFreeCOMSrvに接続するCOMクラスの定義
 */
public class MyRegFreeCOMSrv extends COMLateBindingObject implements AutoCloseable {

	/**
	 * MyRegFreeCOMSrvのコネクションポイントのDIID
	 */
	public static IID DIID_IMyRegFreeCOMSrvEvents = new IID("8C11D374-E2BF-4DEF-89AB-81756137C1D0");

	/**
	 * Nameプロパティの変更前イベント
	 */
	public static class NamePropertyChangingEvent {

		private String name;

		private boolean cancel;

		public String getName() {
			return name;
		}

		public boolean isCancel() {
			return cancel;
		}

		public void setName(String name) {
			this.name = name;
		}

		public void setCancel(boolean cancel) {
			this.cancel = cancel;
		}

		@Override
		public String toString() {
			return "name=" + name + ", cancel=" + cancel;
		}
	}

	/**
	 * Nameプロパティの変更後イベント
	 */
	public static class NamePropertyChangedEvent {

		private String name;

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		@Override
		public String toString() {
			return "name=" + name;
		}
	}

	/**
	 * Javaとしてイベントを受け取るインターフェイス
	 */
	public interface MyRegFreeCOMSrvEventListener {

		/**
		 * 名前が変更される前のイベント
		 * @param evt
		 */
		void namePropertyChanging(NamePropertyChangingEvent evt);

		/**
		 * 名前が変更された後のイベント
		 * @param evt
		 */
		void namePropertyChanged(NamePropertyChangedEvent evt);
	}

	/**
	 * イベントシンクの実装例
	 *
	 * https://github.com/java-native-access/jna/blob/master/contrib/platform/test/com/sun/jna/platform/win32/COM/ComEventCallbacks_Test.java
	 */
	public static class MyRegFreeCOMSrvEventsSink implements IDispatchCallback {

		/**
		 * DISPID(1) NamePropertyChangingイベントのDispId
		 */
		private static final int DISPID_NAME_PROPERTY_CHANGING = 1;

		/**
		 * DISPID(2) NamePropertyChangedイベントのDispId
		 */
		private static final int DISPID_NAME_PROPERTY_CHANGED = 2;

		//------------------------ イベントリスナ ------------------------------

		/**
		 * Javaイベントリスナを保持するリスト
		 */
		private final ConcurrentLinkedDeque<MyRegFreeCOMSrvEventListener>
			eventListeners = new ConcurrentLinkedDeque<>();

		public void addListener(MyRegFreeCOMSrvEventListener listener) {
			eventListeners.add(listener);
		}

		public void removeListener(MyRegFreeCOMSrvEventListener listener) {
			eventListeners.remove(listener);
		}

		//------------------------ JNA ------------------------------

		/**
		 * IDispatchのvtblを作成してCOMからの呼び出しを、
		 * このクラス(IDispatchCallback)に転送できるようにするためのJNAの仕掛け
		 */
		private final DispatchListener listener = new DispatchListener(this);

		@Override
		public Pointer getPointer() {
			return this.listener.getPointer();
		}

		//------------------------ IDispatch ------------------------------

		@Override
		public HRESULT GetTypeInfoCount(UINTByReference pctinfo) {
			return new HRESULT(WinError.E_NOTIMPL); // イベントシンクでは使われないので実装不要
		}

		@Override
		public HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, PointerByReference ppTInfo) {
			return new HRESULT(WinError.E_NOTIMPL); // イベントシンクでは使われないので実装不要
		}

		@Override
		public HRESULT GetIDsOfNames(REFIID riid, WString[] rgszNames, int cNames, LCID lcid,
				DISPIDByReference rgDispId) {
			return new HRESULT(WinError.E_NOTIMPL); // イベントシンクでは使われないので実装不要
		}

		@Override
		public HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
	            WORD wFlags, DISPPARAMS.ByReference pDispParams,
	            VARIANT.ByReference pVarResult, EXCEPINFO.ByReference pExcepInfo,
	            IntByReference puArgErr) {
			System.out.println("★dispIdMember=" + dispIdMember);
			VARIANT[] arguments = pDispParams.getArgs();
			switch (dispIdMember.intValue()) {
	            case DISPID_NAME_PROPERTY_CHANGING: {
	                // 引数は逆順で積まれるので最後の引数が最初にくる
	                VARIANT pCancel = arguments[0];
	                VARIANT name = arguments[1];
	                VARIANT_BOOLByReference cancel = ((VARIANT_BOOLByReference) pCancel.getValue());

	                NamePropertyChangingEvent evt = new NamePropertyChangingEvent();
	                evt.setName(name.stringValue());
	                evt.setCancel(cancel.getValue().booleanValue());

	                for (MyRegFreeCOMSrvEventListener l : eventListeners) {
	                	l.namePropertyChanging(evt);
	                }

	                // byref cancel as boolean の返却
                	cancel.setValue(evt.isCancel() ? Variant.VARIANT_TRUE : Variant.VARIANT_FALSE);
	                break;
	            }

	            case DISPID_NAME_PROPERTY_CHANGED: {
	                VARIANT name = arguments[0];

	                NamePropertyChangedEvent evt = new NamePropertyChangedEvent();
	                evt.setName(name.stringValue());

	                for (MyRegFreeCOMSrvEventListener l : eventListeners) {
	                	l.namePropertyChanged(evt);
	                }
	                break;
	            }
	        }
			return WinError.S_OK;
		}

		//------------------------ IUnknown ------------------------------

		@Override
		public HRESULT QueryInterface(REFIID refiid, PointerByReference ppvObject) {
			if (null == ppvObject) {
                return new HRESULT(WinError.E_POINTER);
            }

            if (refiid.getValue().equals(DIID_IMyRegFreeCOMSrvEvents)) {
                ppvObject.setValue(this.getPointer());
                return WinError.S_OK;
            }

            if (refiid.getValue().equals(IUnknown.IID_IUNKNOWN)) {
                ppvObject.setValue(this.getPointer());
                return WinError.S_OK;
            }

            if (refiid.getValue().equals(IDispatch.IID_IDISPATCH)) {
                ppvObject.setValue(this.getPointer());
                return WinError.S_OK;
            }

            ppvObject.setValue(Pointer.NULL);
            return new HRESULT(WinError.E_NOINTERFACE);
        }

		@Override
		public int AddRef() {
			return 2;
		}

		@Override
		public int Release() {
			return 1;
		}
	}

	/**
	 * イベントシンク。
	 * このオブジェクトと寿命をともにする。
	 * COMからのイベントを受け取ってJavaのインターフェイスにアダプタする。
	 */
	private final MyRegFreeCOMSrvEventsSink eventSink = new MyRegFreeCOMSrvEventsSink();

	/**
	 * コンストラクタ
	 */
	public MyRegFreeCOMSrv() {
		super("MyRegFreeCOMSrv", false);
		try {
			connect();

		} catch (RuntimeException ex) {
			close();
			throw ex;
		}
	}

	/**
	 * リリース
	 */
	@Override
	public void release() {
		disconnect();
		super.release();
	}

	/**
	 * コネクションポイント
	 */
	private ConnectionPoint connectionPoint;

	/**
	 * コネクションポイントにAdviseしたイベントシンクを示すCookie
	 */
	private DWORD cookie;

	/**
	 * コネクションポイントに接続する
	 */
	private void connect() {
		IDispatch pDisp = getIDispatch();

		// コネクションポイントコンテナの取得
		PointerByReference ppCpc = new PointerByReference();
		HRESULT hr = pDisp.QueryInterface(
				new REFIID(ConnectionPointContainer.IID_IConnectionPointContainer), ppCpc);
		COMUtils.checkRC(hr);
		ConnectionPointContainer cpc = new ConnectionPointContainer(ppCpc.getValue());
		try {
			// コネクションポイントの取得
			PointerByReference ppCP = new PointerByReference();
			hr = cpc.FindConnectionPoint(new REFIID(DIID_IMyRegFreeCOMSrvEvents.getPointer()), ppCP);
			COMUtils.checkRC(hr);
			connectionPoint = new ConnectionPoint(ppCP.getValue());
	        DWORDByReference pdwCookie = new DWORDByReference();

	        // 接続する
	        COMUtils.checkRC(connectionPoint.Advise(eventSink, pdwCookie));
	        cookie = pdwCookie.getValue();

		} finally {
			cpc.Release();
		}
	}

	/**
	 * コネクションポイントの接続解除
	 */
	private void disconnect() {
		if (connectionPoint != null) {
			// コネクションポイント解除
			connectionPoint.Unadvise(cookie);
			connectionPoint.Release();
			connectionPoint = null;
		}
	}

	// ------- イベントリスナの追加・削除 -------

	public void addListener(MyRegFreeCOMSrvEventListener l) {
		eventSink.addListener(l);
	}

	public void removeListener(MyRegFreeCOMSrvEventListener l) {
		eventSink.removeListener(l);
	}

	// ------- COMプロパティ・メソッドの呼び出し -------

	public String getName() {
		return super.getStringProperty("Name");
	}

	public void setName(String name) {
		super.setProperty("Name", name);
	}

	public void ShowHello() {
		super.invokeNoReply("ShowHello");
	}

	@Override
	public void close() {
		release();
	}
}