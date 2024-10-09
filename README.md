# GAS-project

Google App Scriptを触ってみたいと思ったので、息抜きに作ってみました。

研究用に調査中の論文をスプレッドシート上で管理するためのGASです。

これから研究が本格化していき、論文のデータが増えた際に、スプレッドシートで管理できたら楽かな？と思い作ってみました。

こんな感じになります。
"開く"を押すとファイルを開くことができます。

![使用イメージ](images/GASイメージ.png)

[使用方法]
1.記入したいスプレッドシートの拡張機能AppScriptにコードをペースト。
2.スプレッドシートのIDと論文を保管しているフォルダのIDを定数SPREADSHEET_IDとPAPERS_FOLDER_IDそれぞれに記入。
3.保存し、onOpenメソッドを実行。
4.スプレッドシート画面に戻り、メニュー欄に追加された'PDF管理'の'PDFをスキャン'をクリック

この一連の流れで使用できます。初めて使用する際は許可を求められます。2回目以降は4番目の作業だけで大丈夫です、
