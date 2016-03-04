InportToSpreadsheet
======================

API経由でGoogleSpreadsheetに書き込む用のサンプルです。  
BatchRequestで大量のデータでも使用できるようにしています。  
https://developers.google.com/google-apps/spreadsheets/?csw=1#SendingBatchRequests


段落を分けるには、[空行](http://example.com/) を入れます。
 
使い方
------
1. [Google Developers Console](https://console.developers.google.com/)から認証情報を作成し、
App.config内に記述します。
2. デスクトップにdata.csvを配置します。
3. Googleドライブ内にスプレッドシートを作成し、Idをソース内に記入します。
4. 実行します。

パラメータの解説
----------------

+ ClientId  
+ ClientSecret  

Developers Consoleで取得します。  
Use Google APIs → Credentials → Create Credentials → OAuth client ID → otherで発行されます。

+ googleAccount

自分のGmailアカウントです。
