# JNARegFreeCOMExample

## 概要

RegFreeCOMExampleで実装した、レジストリフリーのCOMオブジェクトをJAVAから利用する実装例。

- JNAを使いCOMの呼び出しを行う
  - コネクションポイントによるイベントの受け取りも実装している。
- JNAを使い、Win32のアクティベーションコンテキストを明示的に作成して、プログラム実行中にマニフェストによるSide by Side AssemblyによるレジストリフリーのCOM探索を行うようにしている。

