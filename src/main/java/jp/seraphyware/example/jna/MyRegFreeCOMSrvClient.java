package jp.seraphyware.example.jna;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.UncheckedIOException;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;

import com.sun.jna.Platform;
import com.sun.jna.platform.win32.Ole32;

import jp.seraphyware.example.jna.MyRegFreeCOMSrv.MyRegFreeCOMSrvEventListener;
import jp.seraphyware.example.jna.MyRegFreeCOMSrv.NamePropertyChangedEvent;
import jp.seraphyware.example.jna.MyRegFreeCOMSrv.NamePropertyChangingEvent;

/**
 * SxSでレジストリフリーでMyRegFreeCOMSrv COMオブジェクトを起動し利用する実装例。
 *
 * アクティベーションコンテキストの構成には、Win32 APIをJNAから起動している。
 * COMオブジェクトの生成と操作にはJNAのCOMLateBindingObjectを利用している。
 */
public class MyRegFreeCOMSrvClient {

	private static final String CLIENT_MANIFEST = "client.manifest";

	public static void main(String[] args) throws Exception {
		Ole32.INSTANCE.CoInitialize(null);

		// ネイティブのDLLおよびマニフェストファイルをテンポラリに展開する.
		// (DLL、マニフェストともに実ファイルが必要なため)
		String prefix;
		if (Platform.is64Bit()) {
			prefix = "x64";
		} else {
			prefix = "x86";
		}

		String[] resources = {
				"native/" + CLIENT_MANIFEST,
				"native/" + prefix + "/MyRegFreeCOMSrv.dll"
				};

		Path tempDir = Paths.get(System.getProperty("java.io.tmpdir"));
		Path nativeDir = tempDir.resolve(MyRegFreeCOMSrvClient.class.getName())
				.resolve("native").resolve(prefix);
		Files.createDirectories(nativeDir);

		ClassLoader clsldr = MyRegFreeCOMSrvClient.class.getClassLoader();
		for (String resource : resources) {
			try (InputStream is = clsldr.getResourceAsStream(resource)) {
				String fileName = Paths.get(resource).getFileName().toString();
				Path destFile = nativeDir.resolve(fileName);
				try (FileChannel writer = (FileChannel) Files.newByteChannel(destFile,
						StandardOpenOption.WRITE, StandardOpenOption.TRUNCATE_EXISTING, StandardOpenOption.CREATE)) {
					while (writer.transferFrom(Channels.newChannel(is), writer.size(), 4096) > 0);
				}
			}
		}

		// マニフェストファイルの位置
		String manifestFile = nativeDir.resolve(CLIENT_MANIFEST).toString();
		System.out.println("Manifest=" + manifestFile);

		// アクティベーションコンテキストで明示的にマニフェストを読み込んでSxSでレジストリフリーでCOMを構築する
		try (MyRegFreeCOMSrv srv = ActivationContextAPI.doActivate(
				manifestFile, () -> new MyRegFreeCOMSrv())) {
			// 構築されたCOMに対する操作を行う。

			srv.addListener(new MyRegFreeCOMSrvEventListener() {

				@Override
				public void namePropertyChanging(NamePropertyChangingEvent evt) {
	                System.out.println("NamePropertyChanging: " + evt);

	                System.out.println("cancel? (yes/no)");
	                try {
	                	BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
	                	String line;
	                	while ((line = br.readLine()) != null) {
	                		if ("yes".equals(line)) {
	                			evt.setCancel(true);
	                			break;

	                		} else if ("no".equals(line)) {
	                			break;
	                		}
	                	}
	                } catch (IOException ex) {
	                	throw new UncheckedIOException(ex);
	                }
				}

				@Override
				public void namePropertyChanged(NamePropertyChangedEvent evt) {
					System.out.println("Name Changed: " + evt.getName());
				}
			});

			srv.setName("PiyoPiyo");
			srv.ShowHello();
		}
		System.out.println("Done!");

		Ole32.INSTANCE.CoUninitialize();
	}
}

