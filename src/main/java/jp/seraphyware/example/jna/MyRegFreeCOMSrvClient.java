package jp.seraphyware.example.jna;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.UncheckedIOException;
import java.nio.ByteBuffer;
import java.nio.channels.Channels;
import java.nio.channels.ReadableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.Arrays;

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

	/**
	 * ネイティブのDLLおよびマニフェストファイルをテンポラリに展開する.
	 * (DLL、マニフェストともに実ファイルが必要なため)
	 * (また、COM DLLのアンロードはJNAでは制御が難しいので、
	 * アプリケーション終了時に消すのは諦めている。)
	 * @return 展開されたファイルのあるネイティブファイルの位置
	 * @throws IOException
	 */
	public static Path initDLL() throws IOException {
		String prefix;
		if (Platform.is64Bit()) {
			prefix = "x64";
		} else {
			prefix = "x86";
		}

		// 展開するリソース
		String[] resources = {
				"native/" + CLIENT_MANIFEST,
				"native/" + prefix + "/MyRegFreeCOMSrv.dll"
				};

		// 展開先
		String uniqFolder = MyRegFreeCOMSrvClient.class.getName();
		Path tempDir = Paths.get(System.getProperty("java.io.tmpdir"));
		Path nativeDir = tempDir.resolve(uniqFolder)
				.resolve("native").resolve(prefix);
		Files.createDirectories(nativeDir);

		// リソースを展開する
		ClassLoader clsldr = MyRegFreeCOMSrvClient.class.getClassLoader();
		for (String resource : resources) {
			String fileName = Paths.get(resource).getFileName().toString();
			Path destFile = nativeDir.resolve(fileName);

			boolean exist = Files.exists(destFile);
			if (exist) {
				// すでにファイルがある場合
				byte[] fileHash, resHash;

				// 現在のファイルのハッシュ値を求める
				try (ReadableByteChannel fc = Files.newByteChannel(destFile,
						StandardOpenOption.READ)) {
					fileHash = calcSHA1(fc);
				}

				// リソースのハッシュ値を求める
				try (InputStream is = clsldr.getResourceAsStream(resource);
						ReadableByteChannel rbc = Channels.newChannel(is)) {
					resHash = calcSHA1(rbc);
				}

				// リソースとファイルのハッシュが同一であれば存在するものとする.
				// そうでなければ非存在とみなす。
				exist = Arrays.equals(fileHash, resHash);
			}

			if (!exist) {
				try (InputStream is = clsldr.getResourceAsStream(resource)) {
					System.out.println("create: " + destFile);
					Files.copy(is, destFile);
				}
			}
		}

		return nativeDir;
	}

	/**
	 * ハッシュ値を求める
	 * @param ch
	 * @return
	 * @throws IOException
	 */
	private static byte[] calcSHA1(ReadableByteChannel ch) throws IOException {
		try {
			MessageDigest sha1 = MessageDigest.getInstance("SHA1");
			ByteBuffer buf = ByteBuffer.allocate(4096);
			int rd;
			while ((rd = ch.read(buf)) > 0) {
				sha1.update(buf.array(), 0, rd);
				buf.clear();
			}
			return sha1.digest();

		} catch (NoSuchAlgorithmException ex) {
			throw new RuntimeException(ex);
		}
	}

	/**
	 * エントリポイント.
	 * (標準入力から応答を入力する必要があることに注意。)
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		Ole32.INSTANCE.CoInitialize(null);
		System.out.println("java version=" + System.getProperty("java.version"));

		// リソースに格納されているx86/x64用のDLLを
		// ロードできるようにテンポラリに展開する
		Path nativeDir = initDLL();

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
            		System.out.print(">");
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
	                		System.out.print(">");
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

