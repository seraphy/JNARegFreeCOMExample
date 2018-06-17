# JNARegFreeCOMExample

## 概要

Windows環境下で、RegFreeCOMExampleで実装した、レジストリフリーのCOMオブジェクトをJAVAから利用する実装例。

- JNAの[COMLateBindingObject](https://java-native-access.github.io/jna/4.2.1/com/sun/jna/platform/win32/COM/COMLateBindingObject.html) を使いJavaとCOM間の呼び出しをブリッジしている。
  - コネクションポイントによるCOMからのイベントの受け取りも実装している。
  - イベントシンク側(受け取り側)は、[IDispatchCallback](https://java-native-access.github.io/jna/4.2.1/com/sun/jna/platform/win32/COM/IDispatchCallback.html) を使ってコールバックを実装している。
- JNAを使い、Win32の[アクティベーションコンテキスト](https://msdn.microsoft.com/en-us/library/windows/desktop/aa374153.aspx) を明示的に作成して、プログラム実行中にマニフェストによるSide by Side AssemblyによるレジストリフリーのCOM探索を行うようにしている。

## ビルド方法

EclipseのMavenのプロジェクトなので、EclipseまたはMavenによってビルド可能です。

Windows7(x64/x32), Windows10(x64)の環境下の、Java1.8(x64/x86), Java9, 10(x64)での動作を確認しています。

## 注意点

DLLおよびマニフェストファイルは実ファイルでないとOSによって読み込むことができないため、
テンポラリフォルダ上にDLLとマニフェストを展開しています。

DLLはアンロードしなければ削除できませんが、
COMのDLLのアンロードのタイミングはJNAより制御するのが難しいため、アプリケーション終了時に放置しています。

ただし、起動のたびにファイルが増えないようにテンポラリ上の特定の名前でDLLを作成しており、
且つ、同一の内容でない場合のみファイルを書き換えるようにしています。
(したがって通常は二重起動してもファイルの上書きはなく、2つ以上のインスタンスを立ち上げることができます。)

