using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using System.Web;

namespace DocumentManagement_List
{
    class ClassAttributeSetting
    {
        #region <定数定義>

        /// <summary>
        /// CSVファイルの文字コード
        /// </summary>
        public const string CSV_ENCODING = "shift_jis";             // CSVファイルの文字コード

        /// <summary>
        /// 属性設定CSVファイルに「保存期間」「保存期間（年数）」が無い場合にも対応
        /// </summary>
        public const int CSV_COLUMN_COUNT = 6;                      // 20170717 修正

        /// <summary>
        /// 複数キーワード最大数
        /// </summary>
        public const int MAX_KEYWORD_COUNT = 10;                    // 複数キーワード最大数

        /// <summary>
        /// 属性設定CSVファイルのデータ内の「,」を置換する文字列
        /// </summary>
        public const string DELIMITER = "■DELIMITER■";            // 20171017 追加
        #endregion

        #region <内部変数>

        /// <summary>
        /// 属性設定格納リスト
        /// </summary>
        private List<AttributeSetting> ListAttrSet;

        /// <summary>
        /// CSVバージョン情報
        /// </summary>
        public string strCsvVersion;

        #endregion

        #region <クラス定義>
        /// <summary>
        /// クラス定義
        /// </summary>

        /// <summary>
        /// 属性設定情報(+検索用文字列)
        /// </summary>
        public class AttributeSetting
        {

            /// <summary>
            /// 大分類名
            /// </summary>
            public string LargeClassName { get; set; }

            /// <summary>
            /// 中分類名
            /// </summary>
            public string MiddleClassName { get; set; }

            /// <summary>
            /// 小分類名
            /// </summary>
            public string SmallClassName { get; set; }

            /// <summary>
            /// 文書例
            /// </summary>
            public string DocumentExample { get; set; }

            /// <summary>
            /// SAB秘
            /// </summary>
            public string SecrecyLevel { get; set; }

            /// <summary>
            /// 分類コード
            /// </summary>
            public string ClassCode { get; set; }

            /// <summary>
            /// 検索用大分類名
            /// </summary>
            public string LargeClassSearchName { get; set; }

            /// <summary>
            /// 検索用中分類名
            /// </summary>
            public string MiddleClassSearchName { get; set; }

            /// <summary>
            /// 検索用小分類名
            /// </summary>
            public string SmallClassSearchName { get; set; }

            /// <summary>
            /// 検索用文書例
            /// </summary>
            public string DocumentExampleSearch { get; set; }

            /// <summary>
            /// 保存期間名
            /// </summary>
            public string RetentionPeriodName { get; set; }

            /// <summary>
            /// 保存期間
            /// </summary>
            public string RetentionPeriod { get; set; }
        }
        #endregion

        #region <コンストラクタ>
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ClassAttributeSetting()
        {
            // 属性設定情報作成
            ListAttrSet = new List<AttributeSetting>();
        }
        #endregion

        #region メソッド

        /// <summary>
        /// CSVファイル読み込み
        /// </summary>
        /// <param name="strFilePath">CSVファイルパス</param>
        /// <returns>処理結果 True:正常 False:異常</returns>
        public Boolean readCsvFile(string strFilePath)
        {
            Boolean bResult;

            bResult = false;

            int iCsvDataCnt = 0;
            try
            {
                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Start();
                // csvファイルを開く
                using (StreamReader sr = new StreamReader(strFilePath, Encoding.GetEncoding(CSV_ENCODING)))
                {
                    // 末尾まで繰り返す
                    while (!sr.EndOfStream)
                    {
                        // ファイルから一行読み込む
                        string strLine = sr.ReadLine();

                        // 20171017 追加（データにカンマ「,」を含む場合に対応　A,B,"C,D",E,F…）
                        int serch_index = 0;    // 検索インデックス

                        // 末尾までループ検索
                        while (serch_index >= 0)
                        {
                            // 20171017 追加（データにカンマ「,」を含む場合に対応　A,B,"C,D",E,F…）
                            // 行に含まれる「"」を取得
                            serch_index = strLine.IndexOf("\"" , serch_index);

                            // 「"」を含んでいない場合
                            if (serch_index < 0)
                            {
                                // 処理なし
                            }
                            // 「"」を含んでいる場合
                            else
                            {
                                // 1つ目の「"」
                                int start_wQuotation = serch_index;
                                // 2つ目の「"」
                                int end_wQuotation = strLine.IndexOf("\"", start_wQuotation + 1);
                                // 1つ目の「"」と2つ目の「"」の間の「,」を■DELIMITER■に置換
                                ReplaceDelimiter(ref strLine, start_wQuotation, end_wQuotation , ref serch_index);
                                //// デバッグ用
                                //string str = strLine.Substring(serch_index , 1);
                            }                            
                        }
                        
                        // 読み込んだ一行をカンマ毎に分けて配列に格納する
                        string[] strsCsvData = strLine.Split(',');

                        // 項目数が正常であった場合
                        if (CSV_COLUMN_COUNT <= strsCsvData.Count())
                        {
                            // 20171017 追加（データにカンマ「,」を含む場合に対応　A,B,"C,D",E,F…）
                            // ※■DELIMITER■に置換した「,」を元に戻す
                            // 各データ走査
                            for (int i = 0; i < strsCsvData.Length; i++)
                            {
                                // ■DELIMITER■がある場合
                                if (strsCsvData[i].IndexOf(DELIMITER) >= 0)
                                {
                                    // ■DELIMITER■を「,」を元に戻す
                                    strsCsvData[i] = strsCsvData[i].Replace(DELIMITER, ",");
                                    // 不要な「"」を削除
                                    strsCsvData[i] = strsCsvData[i].Replace("\"","");
                                }
                            }

                            if (iCsvDataCnt == 0)
                            {
                                // ヘッダ項目の場合

                                // バージョン情報取得
                                string[] strsVersion = strsCsvData[5].Split(';');
                                strCsvVersion = (strsVersion.Count() == 2) ? strsVersion[1] : "";
                            }
                            else
                            {
                                // データ項目の場合

                                // 検索用文字列作成(大分類、中分類、小分類、文書例)
                                string strLSearchName = getSaerchString(strsCsvData[0].Trim());
                                string strMSearchName = getSaerchString(strsCsvData[1].Trim());
                                string strSSearchName = getSaerchString(strsCsvData[2].Trim());
                                string strDocumentExampleSearch = getSaerchString(strsCsvData[3].Trim());

                                // 20170717 追加（属性設定CSVファイルの「保存期間」「保存期間（年数）」保存用）
                                string retentionPeriodName = "";
                                string retentionPeriod = "";

                                // 20170717 追加（属性設定CSVファイルに「保存期間」「保存期間（年数）」が無い場合）
                                if (CSV_COLUMN_COUNT == strsCsvData.Count())
                                {
                                    //string retentionPeriodName = "";
                                    //string retentionPeriod = "";
                                }
                                // 20170717 修正（属性設定CSVファイルに「保存期間」「保存期間（年数）」がある場合）
                                else
                                {
                                    retentionPeriodName = strsCsvData[6].Trim();
                                    retentionPeriod = strsCsvData[7].Trim();
                                }

                                // 20170717 修正（属性設定CSVファイルに「保存期間」「保存期間（年数）」がある場合、無い場合共通）
                                // 属性設定格納リストに追加
                                ListAttrSet.Add(new AttributeSetting
                                {
                                    LargeClassName = strsCsvData[0].Trim(),
                                    MiddleClassName = strsCsvData[1].Trim(),
                                    SmallClassName = strsCsvData[2].Trim(),
                                    DocumentExample = strsCsvData[3].Trim(),
                                    SecrecyLevel = strsCsvData[4].Trim(),
                                    ClassCode = strsCsvData[5].Trim(),
                                    LargeClassSearchName = strLSearchName,
                                    MiddleClassSearchName = strMSearchName,
                                    SmallClassSearchName = strSSearchName,
                                    DocumentExampleSearch = strDocumentExampleSearch,
                                    RetentionPeriodName = retentionPeriodName,
                                    RetentionPeriod = retentionPeriod
                                });

                            }
                            iCsvDataCnt++;
                        }
                    }
                }

                sw.Stop();
                Console.WriteLine(sw.Elapsed);


                bResult = true;
            }
            catch (Exception ex)
            {
                // ファイルを開くのに失敗したとき
                System.Console.WriteLine(ex.Message);
            }

            return bResult;
        }

        /// <summary>
        /// 属性設定取得
        /// </summary>
        /// <param name="strLargeClassName">分類コード</param>
        /// <returns></returns>
        public IEnumerable<AttributeSetting> getAttributeSettingList(string strClassCode)
        {
            // 検索処理
            IEnumerable<AttributeSetting> ListRet;
            ListRet =
                from las in ListAttrSet
                where las.ClassCode.Contains(strClassCode)
                orderby las.LargeClassName, las.MiddleClassName, las.SmallClassName, las.DocumentExample
                select las;
            return ListRet;
        }

        /// <summary>
        /// 属性設定取得
        /// </summary>
        /// <param name="strLargeClassName">検索大分類名称</param>
        /// <param name="strMiddleClassName">検索中分類名称</param>
        /// <param name="strSmallClassName">検索小分類名称</param>
        /// <param name="strDocumentExample">検索文書例名称</param>
        /// <returns></returns>
        public IEnumerable<AttributeSetting> getAttributeSettingList(string strLargeClassName, string strMiddleClassName, string strSmallClassName, string strDocumentExample)
        {
            // 検索処理
            IEnumerable<AttributeSetting> ListRet;
            ListRet =
                from las in ListAttrSet
                where las.LargeClassName.Contains(strLargeClassName) &&
                      las.MiddleClassName.Contains(strMiddleClassName) &&
                      las.SmallClassName.Contains(strSmallClassName) &&
                      las.DocumentExample.Contains(strDocumentExample)
                orderby las.LargeClassName, las.MiddleClassName, las.SmallClassName, las.DocumentExample
                select las;
            return ListRet;
        }

        /// <summary>
        /// キーワード検索
        /// </summary>
        /// <param name="strSearchWord">検索文字列</param>
        /// <returns>検索結果リスト</returns>
        public IEnumerable<AttributeSetting> SearchList(string strSearchWord)
        {
            #region 1ワード検索処理
            /*
             * 1ワード検索処理 2016/06/01
            // 検索文字列で不要なスペースは削除する
            strSearchWord = strSearchWord.Trim();

            // 検索処理
            IEnumerable<AttributeSetting> ListRet;
            ListRet =
                from las in ListAttrSet
                where las.LargeClassSearchName.Contains(strSearchWord) ||
                      las.MiddleClassSearchName.Contains(strSearchWord) ||
                      las.SmallClassSearchName.Contains(strSearchWord) ||
                      las.DocumentExampleSearch.Contains(strSearchWord) ||
                      las.SecrecyLevel.Contains(strSearchWord) ||
                      las.ClassCode.Contains(strSearchWord)
                orderby las.LargeClassName, las.MiddleClassName, las.SmallClassName, las.DocumentExample
                select las;

            return ListRet;
            */
            #endregion

            #region 複数ワード処理
            // 複数ワード検索処理
            string[] stSearchWord;

            // 入力キーワードを分割し検索ワード作成
            strSearchWord = strSearchWord.Replace("　", " ");
            stSearchWord = strSearchWord.Split(' ');

            // 入力キーワードを検索キーワードに変換
            for (int i = 0; i < stSearchWord.Count(); i++)
            {
                stSearchWord[i] = getSaerchString(stSearchWord[i]);
            }

            // 全データを一旦取得
            IEnumerable<AttributeSetting> ListRet = from las in ListAttrSet select las;
            return SearchListReProc(ListRet, stSearchWord, 0);
            #endregion
        }


        /// <summary>
        /// 複数キーワード検索(再起処理)
        /// </summary>
        /// <param name="ListData">検索結果リスト</param>
        /// <param name="stSearchWord">検索文字列リスト</param>
        /// <param name="iSearchWordNo">検索文字列番号</param>
        /// <returns></returns>
        public IEnumerable<AttributeSetting> SearchListReProc(IEnumerable<AttributeSetting> ListData ,string[] stSearchWord, int iSearchWordNo)
        {
            // 検索データ終了時
            if (stSearchWord.Count() <= iSearchWordNo)
            {
                return ListData;
            }

            // 検索処理
            ListData =
                from las in ListData
                where
                    las.LargeClassSearchName.Contains(stSearchWord[iSearchWordNo].Trim()) ||
                    las.MiddleClassSearchName.Contains(stSearchWord[iSearchWordNo].Trim()) ||
                    las.SmallClassSearchName.Contains(stSearchWord[iSearchWordNo].Trim()) ||
                    las.DocumentExampleSearch.Contains(stSearchWord[iSearchWordNo].Trim()) ||
                    las.SecrecyLevel.Contains(stSearchWord[iSearchWordNo].Trim()) ||
                    las.ClassCode.Contains(stSearchWord[iSearchWordNo].Trim())
                orderby las.LargeClassName, las.MiddleClassName, las.SmallClassName, las.DocumentExample
                select las;

            // 検索結果を新規リストに登録
            List<AttributeSetting> ListRet = new List<AttributeSetting>();
            foreach (ClassAttributeSetting.AttributeSetting row in ListData)
            {
                ListRet.Add(new AttributeSetting
                {
                    LargeClassName = row.LargeClassName.Trim(),
                    MiddleClassName = row.MiddleClassName.Trim(),
                    SmallClassName = row.SmallClassName.Trim(),
                    DocumentExample = row.DocumentExample.Trim(),
                    SecrecyLevel = row.SecrecyLevel.Trim(),
                    ClassCode = row.ClassCode.Trim(),
                    LargeClassSearchName = row.LargeClassSearchName.Trim(),
                    MiddleClassSearchName = row.MiddleClassSearchName.Trim(),
                    SmallClassSearchName = row.SmallClassSearchName.Trim(),
                    DocumentExampleSearch = row.DocumentExampleSearch.Trim()
                });
            }

            // 次の検索へ
            iSearchWordNo++;
            return SearchListReProc(ListRet, stSearchWord, iSearchWordNo);
        }

        /// <summary>
        /// 検索用文字列の生成処理
        /// </summary>
        /// <param name="strBefore"></param>
        /// <returns>strAfter</returns>
        public string getSaerchString(string strBefore)
        {
            string strAfter;

            // 全角スペースの削除
            strAfter = strBefore.Replace("　", "");

            // 伸ばし棒を「ー」に統一変換する
            strAfter = strAfter.Replace("ｰ", "ー");
            strAfter = strAfter.Replace("-", "ー");
            strAfter = strAfter.Replace("－", "ー");
            strAfter = strAfter.Replace("─", "ー");

            // 全角を半角に変換
            strAfter = Strings.StrConv(strAfter, VbStrConv.Narrow, 0x411); // 2016/09/20 ロケール設定を追加

            // 小文字を大文字に変換
            strAfter = strAfter.ToUpper();

            return strAfter;
        }

        /// <summary>
        /// 大分類一覧取得
        /// </summary>
        /// <returns></returns>
        public string[] getLargeClassNameList()
        {
            // 大分類絞込み処理
            var ListRet = 
                from las in ListAttrSet
                group las by las.LargeClassName into g
	            select new
	            {
                    LargeClassName = g.Key
	            };

            // 処理結果セット
            int iRowCnt = 0;
            string[] strList = new string[ListRet.Count()];
            foreach (var row in ListRet)
            {
                strList[iRowCnt] = row.LargeClassName.ToString();
                iRowCnt++;
            }

            return strList;
        }

        /// <summary>
        /// 中分類一覧取得
        /// </summary>
        /// <param name="strLargeClassName">大分類名</param>
        /// <returns></returns>
        public string[] getMiddleClassNameList(string strLargeClassName)
        {
            // 中分類絞込み処理
            var ListRet =
                from las in ListAttrSet
                where las.LargeClassName == strLargeClassName
                group las by las.MiddleClassName into g
                select new
                {
                    MiddleClassName = g.Key
                };

            // 処理結果セット
            int iRowCnt = 0;
            string[] strList = new string[ListRet.Count()];
            foreach (var row in ListRet)
            {
                strList[iRowCnt] = row.MiddleClassName.ToString();
                iRowCnt++;
            }

            return strList;
        }

        /// <summary>
        /// 小分類一覧取得
        /// </summary>
        /// <param name="strLargeClassName">大分類名</param>
        /// <param name="strMiddleClassName">中分類名</param>
        /// <returns></returns>
        public string[] getSmallClassNameList(string strLargeClassName, string strMiddleClassName)
        {
            // 小分類絞込み処理
            var ListRet =
                from las in ListAttrSet
                where las.LargeClassName == strLargeClassName && las.MiddleClassName == strMiddleClassName
                group las by las.SmallClassName into g
                select new
                {
                    SmallClassName = g.Key
                };

            // 処理結果セット
            int iRowCnt = 0;
            string[] strList = new string[ListRet.Count()];
            foreach (var row in ListRet)
            {
                strList[iRowCnt] = row.SmallClassName.ToString();
                iRowCnt++;
            }

            return strList;
        }

        /// <summary>
        /// 文書例一覧取得
        /// </summary>
        /// <param name="strLargeClassName"></param>
        /// <param name="strMiddleClassName"></param>
        /// <param name="strSmallClassName"></param>
        /// <returns></returns>
        public string[] getDocumentExampleList(string strLargeClassName, string strMiddleClassName, string strSmallClassName)
        {
            // 文書例絞込み処理
            var ListRet =
                from las in ListAttrSet
                where las.LargeClassName == strLargeClassName && las.MiddleClassName == strMiddleClassName && las.SmallClassName == strSmallClassName
                group las by las.DocumentExample into g
                select new
                {
                    DocumentExample = g.Key
                };

            // 処理結果セット
            int iRowCnt = 0;
            string[] strList = new string[ListRet.Count()];
            foreach (var row in ListRet)
            {
                strList[iRowCnt] = row.DocumentExample.ToString();
                iRowCnt++;
            }

            return strList;
        }

        /// <summary>
        /// 検索文字列内の「,」をDELIMITERに置換する関数
        /// </summary>
        /// <param name="serch_str">検索文字列</param>
        /// <param name="start_str">検索開始番号</param>
        /// <param name="end_str">検索終了番号</param>
        /// <param name="serch_index">検索インデックス</param>
        public void ReplaceDelimiter(ref string serch_str , int start_str , int end_str , ref int serch_index)
        {
            // 1つ目の「"」と2つ目の「"」の間の1データを取得
            string data = serch_str.Substring(start_str + 1, end_str - start_str - 1);
            // 1つ目の「"」と2つ目の「"」の間の「,」を■DELIMITER■に置換
            string delimiter_data = data.Replace(",", DELIMITER);
            // 置換前の検索文字列の長さを保存
            int before_str_len = serch_str.Length;
            // 読み込んだ基データに反映
            serch_str = serch_str.Replace(data, delimiter_data);
            // 置換後の検索文字列の長さを保存
            int after_str_len = serch_str.Length;
            // 検索インデックス更新（2つ目の「"」+ DELIMITERで置き換えた分の増加分 +１（その1つ後ろ））
            serch_index = end_str + (after_str_len - before_str_len) + 1;
            // ※■DELIMITER■に置換した「,」は後で元に戻す
        }

        #endregion
    }
}
