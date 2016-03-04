using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;

namespace ImportToSpreadSheet
{
    class Program
    {
        #region private class

        private class CellAddress
        {
            public uint Row;
            public uint Col;
            public string IdString;

            /**
             * Constructs a CellAddress representing the specified {@code row} and
             * {@code col}. The IdString will be set in 'RnCn' notation.
             */
            public CellAddress(uint row, uint col)
            {
                this.Row = row;
                this.Col = col;
                this.IdString = string.Format("R{0}C{1}", row, col);
            }
        }

        #endregion

        static void Main(string[] args)
        {
            // デスクトップのdata.csvを対象
            var path = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\data.csv";
            Console.WriteLine(path);

            // ファイルの中身を読み取り
            var lines = File.ReadAllLines(path, Encoding.GetEncoding("shift-jis"));
            var sheetId = "";

            var service = GetServiceAsServiceAccount();
            //var service = GetServiceAsOAuth();

            ImportToSpreadsheet(lines, sheetId, service);

            Console.WriteLine("終了するにはなにかキーを押してください。。。");
            Console.Read();
        }

        #region static method

        /// <summary>
        /// OAuth認証で使用するserviceを取得
        /// </summary>
        /// <returns></returns>
        private static SpreadsheetsService GetServiceAsOAuth()
        {
            // 認証情報
            var userName = ConfigurationManager.AppSettings["googleAccountName"];
            var clientId = ConfigurationManager.AppSettings["googleClientId"];
            var clientSecret = ConfigurationManager.AppSettings["googleClientSecret"];

            var userCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = clientId,
                    ClientSecret = clientSecret
                },
                new[]
                {
                    "https://spreadsheets.google.com/feeds"
                },
                userName,
                CancellationToken.None,
                new FileDataStore("myDataStore")
                ).Result;

            var auth = new OAuth2Parameters
            {
                ClientId = clientId,
                ClientSecret = clientSecret,
                RedirectUri = "urn:ietf:wg:oauth:2.0:oob",
                Scope = "https://spreadsheets.google.com/feeds",
                AccessToken = userCredential.Token.AccessToken,
                RefreshToken = userCredential.Token.RefreshToken,
                TokenType = userCredential.Token.TokenType,
            };

            var requestFactory = new GOAuth2RequestFactory(null, "myApp", auth);
            var service = new SpreadsheetsService("myApp")
            {
                Credentials = new GDataCredentials(userCredential.Token.TokenType + " " + userCredential.Token.AccessToken),
                RequestFactory = requestFactory
            };

            return service;
        }

        /// <summary>
        /// ServiceAccountで使用するserviceを取得
        /// </summary>
        /// <returns></returns>
        private static SpreadsheetsService GetServiceAsServiceAccount()
        {
            // 鍵ファイル
            var keyFilePath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\key.p12";

            var serviceAccountEmail = ConfigurationManager.AppSettings["serviceAccountEmail"].ToString();   // found in developer console
            var certificate = new X509Certificate2(keyFilePath, "notasecret", X509KeyStorageFlags.Exportable);

            ServiceAccountCredential credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccountEmail) //create credential using certigicate
            {
                Scopes = new[] { "https://spreadsheets.google.com/feeds/" } //this scopr is for spreadsheets, check google scope FAQ for others
            }.FromCertificate(certificate));

            credential.RequestAccessTokenAsync(System.Threading.CancellationToken.None).Wait(); //request token

            var requestFactory = new GDataRequestFactory("My App User Agent");
            requestFactory.CustomHeaders.Add(string.Format("Authorization: Bearer {0}", credential.Token.AccessToken));

            var service = new SpreadsheetsService("myApp")
            {
                RequestFactory = requestFactory
            };

            return service;
        }

        /// <summary>
        /// スプレッドシートに記入する
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="sheetId"></param>
        /// <param name="service"></param>
        private static void ImportToSpreadsheet(string[] lines, string sheetId, SpreadsheetsService service)
        {
            // 更新対象のスプレッドシートを検索
            // スプレッドシートのId。URLのランダムな文字列の部分            
            var url = "https://spreadsheets.google.com/feeds/spreadsheets/" + sheetId;
            var query = new SpreadsheetQuery(url);
            var feed = service.Query(query);
            var spreadsheet = (SpreadsheetEntry)feed.Entries.First();

            // シートを追加
            var sheetName = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
            // csvの1行目を使って列数を設定
            var colCount = lines[0].Split(',').Count();
            // 行数
            var rowCount = lines.Count();
            var worksheet = service.Insert(spreadsheet.Worksheets, new WorksheetEntry((uint)rowCount, (uint)colCount, sheetName));

            var cellQuery = new CellQuery(worksheet.CellFeedLink);
            var cellFeed = service.Query(cellQuery);


            // セルのアドレスを格納
            var cellAddresses = new List<CellAddress>(rowCount * colCount);
            for (uint row = 1; row <= rowCount; ++row)
            {
                for (uint col = 1; col <= colCount; ++col)
                {
                    cellAddresses.Add(new CellAddress(row, col));
                }
            }

            // バッチ処理
            var cellEntries = GetCellEntryMap(service, cellFeed, cellAddresses);
            var batchRequest = new CellFeed(cellQuery.Uri, service);

            // セルに書き込む値を設定
            for (var i = 0; i < rowCount; i++)
            {
                var line = lines[i];
                var values = line.Split(',');
                for (var j = 0; j < colCount; j++)
                {
                    var cellAddr = cellAddresses[i * colCount + j];
                    var batchEntry = cellEntries[cellAddr.IdString];
                    batchEntry.InputValue = values[j];
                    batchEntry.BatchData = new GDataBatchEntryData(cellAddr.IdString, GDataBatchOperationType.update);
                    batchRequest.Entries.Add(batchEntry);
                }

            }

            CellFeed batchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));
            Console.WriteLine("finish");
        }
        /**
         * Connects to the specified {@link SpreadsheetsService} and uses a batch
         * request to retrieve a {@link CellEntry} for each cell enumerated in {@code
         * cellAddrs}. Each cell entry is placed into a map keyed by its RnCn
         * identifier.
         *
         * https://developers.google.com/google-apps/spreadsheets/?csw=1#SendingBatchRequests
         *
         * @param service the spreadsheet service to use.
         * @param cellFeed the cell feed to use.
         * @param cellAddrs list of cell addresses to be retrieved.
         * @return a dictionary consisting of one {@link CellEntry} for each address in {@code
         *         cellAddrs}
         */
        private static Dictionary<String, CellEntry> GetCellEntryMap(
            SpreadsheetsService service, CellFeed cellFeed, List<CellAddress> cellAddrs)
        {
            CellFeed batchRequest = new CellFeed(new Uri(cellFeed.Self), service);
            foreach (CellAddress cellId in cellAddrs)
            {
                CellEntry batchEntry = new CellEntry(cellId.Row, cellId.Col, cellId.IdString);
                batchEntry.Id = new AtomId(string.Format("{0}/{1}", cellFeed.Self, cellId.IdString));
                batchEntry.BatchData = new GDataBatchEntryData(cellId.IdString, GDataBatchOperationType.query);
                batchRequest.Entries.Add(batchEntry);
            }

            CellFeed queryBatchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));

            Dictionary<String, CellEntry> cellEntryMap = new Dictionary<String, CellEntry>();
            foreach (CellEntry entry in queryBatchResponse.Entries)
            {
                cellEntryMap.Add(entry.BatchData.Id, entry);
                //Console.WriteLine("batch {0} (CellEntry: id={1} editLink={2} inputValue={3})",
                //    entry.BatchData.Id, entry.Id, entry.EditUri,
                //    entry.InputValue);
            }

            return cellEntryMap;
        }

        #endregion
    }
}
