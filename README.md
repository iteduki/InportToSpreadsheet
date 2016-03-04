InportToSpreadsheet
======================

API経由でGoogleSpreadsheetに書き込む用のサンプルです。  
BatchRequestで大量のデータでも使用できるようにしています。  
https://developers.google.com/google-apps/spreadsheets/?csw=1#SendingBatchRequests

 
使い方
------
#OAuth認証の場合

1. [Google Developers Console](https://console.developers.google.com/)から認証情報を作成し、
App.config内に記述します。（取得方法はパラメータの解説に記載）
2. デスクトップにdata.csvを配置します。
3. Googleドライブ内にスプレッドシートを作成し、Idをソース内に記入します。
4. 実行します。
 
#ServiceAccountの場合
1. [Google Developers Console](https://console.developers.google.com/)から認証情報を作成し、
p12鍵ファイルを取得しkey.p12としてデスクトップに配置します。
2. serviceAccountを作成し、serviceAccountEmailをApp.config内に記述します。
3. デスクトップにdata.csvを配置します。
4. Googleドライブ内にスプレッドシートを作成し、Idをソース内に記入します。
5. 対象のスプレッドシートにserviceAccountEmailを編集者として設定します。
6. 実行します。


パラメータの解説
----------------

+ ClientId(OAuthの場合のみ)  
+ ClientSecret(OAuthの場合のみ)  

Developers Consoleで取得します。  
Use Google APIs → Credentials → Create Credentials → OAuth client ID → otherで発行されます。

+ googleAccount(OAuthの場合のみ)

スプレッドシートにアクセスするGoogleアカウントです。
スプレッドシートに編集権限を設定する必要があります。

+ serviceAccountEmail(ServiceAccountの場合のみ)
+ key.p12ファイル
Developers Consoleで取得します。  
Permissiong → Service Account → Create service Account で作成。
作成時にjsonかp12ファイルかを選択させられるのでp12を選択すると、アカウント発行後に鍵ファイルがダウンロードされます。
