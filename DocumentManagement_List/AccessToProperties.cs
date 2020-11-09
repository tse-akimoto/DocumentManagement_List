using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using System.Xml;



namespace DocumentManagement_List
{
    class PropertiesKeyList
    {
        public const string STR_TAG_DC = "dc";
        public const string STR_TAG_CP = "cp";
        public const string STR_TAG_DCTRMS = "dcterms";
        public const string STR_TAG_DCMITYPE = "dcmitype";
        public const string STR_TAG_XSI = "xsi";

        public const string STR_TITLE = "dc:title";
        public const string STR_SUBJECT = "dc:subject";
        public const string STR_CREATOR = "dc:creator";
        public const string STR_KEYWORDS = "cp:keywords";
        public const string STR_DESCRIPTION = "dc:description";
        public const string STR_LAST_MODIFIED_BY = "cp:lastModifiedBy";
        public const string STR_REVISION = "cp:revision";
        public const string STR_CREATED = "dcterms:created";
        public const string STR_MODIFIED = "dcterms:modified";
        public const string STR_CATEGORY = "cp:category";
        public const string STR_CONTENT_STATUS = "cp:contentStatus";
        public const string STR_LANGUAGE = "dc:language";
        public const string STR_VERSION = "cp:version";

        public const string STR_CORE_PROPERTIES = "//cp:coreProperties/";

        // OFFICEの保護状態が最終版の時の判定用
        //public const string STR_FINAL_CONTENT = "最終版";    // step2 iwasa

        /// <summary>
        /// プロパティの一覧を返す
        /// </summary>
        /// <returns></returns>
        public static List<string> getPropertiesKeyList()
        {
            List<string> list = new List<string>();

            list.Add(STR_TITLE);
            list.Add(STR_SUBJECT);
            list.Add(STR_CREATOR);
            list.Add(STR_KEYWORDS);
            list.Add(STR_DESCRIPTION);
            list.Add(STR_LAST_MODIFIED_BY);
            list.Add(STR_REVISION);
            list.Add(STR_CREATED);
            list.Add(STR_MODIFIED);
            list.Add(STR_CATEGORY);
            list.Add(STR_CONTENT_STATUS);
            list.Add(STR_LANGUAGE);
            list.Add(STR_VERSION);

            return list;
        }
    }

    class PropertiesSchemaList
    {
        // スキーマ定義
        private const string corePropertiesSchema = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private const string dcPropertiesSchema = "http://purl.org/dc/elements/1.1/";
        private const string dctermsPropertiesSchema = "http://purl.org/dc/terms/";
        private const string dcmitypePropertiesSchema = "http://purl.org/dc/dcmitype/";
        private const string xsiPropertiesSchema = "http://www.w3.org/2001/XMLSchema-instance";

        public static string CorePropertiesSchema
        {
            get
            {
                return corePropertiesSchema;
            }
        }

        public static string DcPropertiesSchema
        {
            get
            {
                return dcPropertiesSchema;
            }
        }

        public static string DctermsPropertiesSchema
        {
            get
            {
                return dctermsPropertiesSchema;
            }
        }

        public static string DcmitypePropertiesSchema
        {
            get
            {
                return dcmitypePropertiesSchema;
            }
        }

        public static string XsiPropertiesSchema
        {
            get
            {
                return xsiPropertiesSchema;
            }
        }
    }

    class AccessToProperties
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public AccessToProperties()
        {

        }

        /// <summary>
        /// プロパティを書き込みます
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="propertiesTable"></param>
        /// <returns></returns>
        public bool WriteProperties(string filePath, string keyword, int filetype, bool bFileInfoUpdate, List<string> lstFinal, ref DateTime LastWriteTime, ref string error_reason)  // 20171011 修正（ファイル書込み時のエラー理由を保存） // step2 iwasa
        {
            bool bRet = true;

            // 20171010 修正（try-catch追加（対象ファイルが存在しない場合に発生するエラー対応））
            // （検索ボタン押下後、ファイルを移動/削除した場合など）
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(filePath);
                if (fi.Length == 0)
                {
                    // ファイルサイズがゼロの場合プロパティが存在しないのでエラー
                    // 20171228 追加（エラー理由保存）
                    error_reason = ListForm.LIST_VIEW_NA;
                    return false;
                }

                // 最終版
            }
            catch
            {
                // 対象ファイルが存在しない場合にここに来る
                // ファイルが存在しない場合のエラー
                error_reason = ListForm.LIST_VIEW_NA;
                return false;
            }
            
            // ファイルのプロパティ領域を開く
            SpreadsheetDocument excel = null;
            WordprocessingDocument word = null;
            PresentationDocument ppt = null;

            CoreFilePropertiesPart coreFileProperties;

            // 20171010 修正（try-catch追加（開いているファイルを開こうとして発生するエラー対応））
            try
            {
                switch (filetype)
                {
                    case ListForm.EXTENSION_EXCEL:
                        excel = SpreadsheetDocument.Open(filePath, true);
                        coreFileProperties = excel.CoreFilePropertiesPart;
                        break;
                    case ListForm.EXTENSION_WORD:
                        word = WordprocessingDocument.Open(filePath, true);
                        coreFileProperties = word.CoreFilePropertiesPart;
                        break;
                    case ListForm.EXTENSION_POWERPOINT:
                        ppt = PresentationDocument.Open(filePath, true);
                        coreFileProperties = ppt.CoreFilePropertiesPart;
                        break;
                    default:
                        // 異常なファイル
                        // 20171228 追加（エラー理由保存）
                        error_reason = ListForm.LIST_VIEW_NA;
                        return false;
                }
            }
            catch
            {
                // 開いているファイルを開いた場合にここに来る
                // ファイルが開かれている場合のエラー
                // または読み取り専用のファイル
                error_reason = ListForm.LIST_VIEW_NA;
                return false;
            }

            LastWriteTime = new DateTime();

            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace(PropertiesKeyList.STR_TAG_CP, PropertiesSchemaList.CorePropertiesSchema);
            nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DC, PropertiesSchemaList.DcPropertiesSchema);
            nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DCTRMS, PropertiesSchemaList.DctermsPropertiesSchema);
            nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DCMITYPE, PropertiesSchemaList.DcmitypePropertiesSchema);
            nsManager.AddNamespace(PropertiesKeyList.STR_TAG_XSI, PropertiesSchemaList.XsiPropertiesSchema);

            XmlDocument xdoc = new XmlDocument();

            try
            {
                xdoc.Load(coreFileProperties.GetStream());

                // 最終版のチェック
                string searchString = string.Format(PropertiesKeyList.STR_CORE_PROPERTIES + "{0}", PropertiesKeyList.STR_CONTENT_STATUS);
                // 書き込み先を検索
                XmlNode xNode = xdoc.SelectSingleNode(searchString, nsManager);

                //if (xNode != null && xNode.InnerText == PropertiesKeyList.STR_FINAL_CONTENT)
                if (xNode != null && lstFinal.Contains(xNode.InnerText))    // step2 iwasa
                {
                    // 最終版のファイルへの書き込みはできないのでエラーを返す
                    // 20171228追加（エラーの理由を保存）
                    error_reason = ListForm.LIST_VIEW_NA;
                    return false;
                }

                // 更新日時を読み込む
                // 書き込み先のキーワードを指定
                searchString = string.Format(PropertiesKeyList.STR_CORE_PROPERTIES + "{0}", PropertiesKeyList.STR_MODIFIED);
                // 書き込み先を検索
                xNode = xdoc.SelectSingleNode(searchString, nsManager);

                if (xNode != null)
                {
                    LastWriteTime = DateTime.Parse(xNode.InnerText);
                }
                else
                {
                    // なかったら今日の日付
                    LastWriteTime = DateTime.Now;
                }

                // キーワードの書込
                // 書き込み先のキーワードを指定
                searchString = string.Format(PropertiesKeyList.STR_CORE_PROPERTIES + "{0}", PropertiesKeyList.STR_KEYWORDS);
                // 書き込み先を検索
                xNode = xdoc.SelectSingleNode(searchString, nsManager);

                if (xNode != null)
                {
                    // 書き込む
                    xNode.InnerText = keyword;
                }
                else
                {
                    // keywordsタグが存在していないので作成する
                    XmlNode node = xdoc.DocumentElement;

                    // TSE kitada Comment.
                    // .NETのバグ(？)により、"cp:keywords"を出力すると「:」より前をprefix扱いして勝手に削除してしまうため
                    // Excelのフォーマットとして破損した形で出力されてしまう。
                    // "cp:"の部分は名前空間を指定することで勝手に付与されるので、qualifiedNameにはcp:をつけなくてよい。
                    XmlElement el = xdoc.CreateElement(PropertiesKeyList.STR_KEYWORDS, PropertiesSchemaList.CorePropertiesSchema);
                    el.InnerText = keyword;
                    node.AppendChild(el);
                }            
                
                // カテゴリのリセット
                // 書き込み先のキーワードを指定
                searchString = string.Format(PropertiesKeyList.STR_CORE_PROPERTIES + "{0}", PropertiesKeyList.STR_CATEGORY);
                // 書き込み先を検索
                xNode = xdoc.SelectSingleNode(searchString, nsManager);

                if (xNode != null)
                {
                    // 書き込む
                    xNode.InnerText = "";
                    // 保存
                    //xdoc.Save(coreFileProperties.GetStream());
                }
                
                // 20190705 TSE matsuo ファイルのプロパティ領域を再作成する
                if (excel != null)
                {
                    excel.DeletePart(excel.CoreFilePropertiesPart);
                    excel.AddCoreFilePropertiesPart();
                    coreFileProperties = excel.CoreFilePropertiesPart;
                }
                if (word != null)
                {
                    word.DeletePart(word.CoreFilePropertiesPart);
                    word.AddCoreFilePropertiesPart();
                    coreFileProperties = word.CoreFilePropertiesPart;
                }
                if (ppt != null)
                {
                    ppt.DeletePart(ppt.CoreFilePropertiesPart);
                    ppt.AddCoreFilePropertiesPart();
                    coreFileProperties = ppt.CoreFilePropertiesPart;
                }

                // 20180109 TSE kitada 保存は最後に一回だけ
                // 保存
                xdoc.Save(coreFileProperties.GetStream());

            }
            catch (Exception e)
            {
                bRet = false;
                // 20171011 追加 （読取専用、ファイ存在しない以外のエラー）
                error_reason = ListForm.LIST_VIEW_NA;
            }
            finally
            {
                // ファイルのプロパティ領域を閉じる
                if (excel != null) excel.Close();
                if (word != null) word.Close();
                if (ppt != null) ppt.Close();
            }
            try
            {
                // 更新日は更新しない場合は元の更新日を上書きする
                if (!bFileInfoUpdate)
                {
                    System.IO.File.SetLastWriteTime(filePath, LastWriteTime);
                }
            }
            catch (Exception e)
            {
                // 対象ファイルがロックされている場合はスルーする
            }

            return bRet;
        }

        /// <summary>
        /// プロパティを読み込みます
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="propertiesTable"></param>
        /// <returns></returns>
        public Dictionary<string, string> ReadProperties(string filePath, int filetype, ref string error_reason, ref string str_exception)  // 20171011 修正（ファイル読込時のエラー理由を保存）
        {
            Dictionary<string, string> retTable = new Dictionary<string, string>();

            System.IO.FileInfo fi = new System.IO.FileInfo(filePath);
            if(fi.Length == 0)
            {
                // ファイルサイズがゼロの場合プロパティが存在しないのでエラー
                // 20171011 追加 （読取専用、ファイ存在しない以外のエラー）
                error_reason = ListForm.LIST_VIEW_NA;
                return retTable;
            }

            SpreadsheetDocument excel = null;
            WordprocessingDocument word = null;
            PresentationDocument ppt = null;

            CoreFilePropertiesPart coreFileProperties;

            // 20171011 修正（xlsxとdocxのみファイルを開いている状態でファイルアクセスするとエラーになる為の回避対応）
            try
            {
                // ファイルのプロパティ領域を開く
                switch (filetype)
                {
                    case ListForm.EXTENSION_EXCEL:
                        excel = SpreadsheetDocument.Open(filePath, false);
                        coreFileProperties = excel.CoreFilePropertiesPart;
                        break;
                    case ListForm.EXTENSION_WORD:
                        word = WordprocessingDocument.Open(filePath, false);
                        coreFileProperties = word.CoreFilePropertiesPart;
                        break;
                    case ListForm.EXTENSION_POWERPOINT:
                        ppt = PresentationDocument.Open(filePath, false);
                        coreFileProperties = ppt.CoreFilePropertiesPart;
                        break;
                    default:
                        // 異常なファイル
                        // 20171228 追加（エラー理由保存）
                        error_reason = ListForm.LIST_VIEW_NA;
                        return retTable;
                }
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_CP, PropertiesSchemaList.CorePropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DC, PropertiesSchemaList.DcPropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DCTRMS, PropertiesSchemaList.DctermsPropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DCMITYPE, PropertiesSchemaList.DcmitypePropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_XSI, PropertiesSchemaList.XsiPropertiesSchema);

                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(coreFileProperties.GetStream());

                // プロパティのキーリストを作成
                List<string> propertieslist = PropertiesKeyList.getPropertiesKeyList();

                // 全キーリストを見て存在するデータを取得
                foreach (string key in propertieslist)
                {
                    // 書き込み先のキーワードを指定
                    string searchString = string.Format(PropertiesKeyList.STR_CORE_PROPERTIES + "{0}", key);
                    // 書き込み先を検索
                    XmlNode xNode = xdoc.SelectSingleNode(searchString, nsManager);

                    if (xNode != null)
                    {
                        // 読み込む
                        retTable.Add(key, xNode.InnerText);
                    }
                }

                // ファイルのプロパティ領域を閉じる
                if( excel != null ) excel.Close();
                if( word  != null ) word.Close();
                if( ppt   != null ) ppt.Close();
            }
#if false
#region HyperLink修復

            // ■ ADD TSE Kitada
            // HyperLinkが破損している場合に、そのリンクを書き直して正常にOPEN出来るようにする。
            // 但し、ドキュメントの中身を直接書き換える処理のため、見送る。
            catch (OpenXmlPackageException ope)
            {
                if (ope.ToString().Contains("Invalid Hyperlink"))
                {
                    // HyperLinkの破損が原因なので、内部のリンクを修正する
                    using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                    }

                    if (count >= 1)
                    {
                        // 2回実行してダメだったので終了
                        error_reason = ListForm.LIST_VIEW_NA;
                    }
                    else
                    {
                        // もう一度トライ
                        retTable = ReadProperties(filePath, filetype, ref error_reason, ref str_exception, 1);
                    }
                }

                if (excel != null) excel.Close();
                if (word != null) word.Close();
                if (ppt != null) ppt.Close();
            }
#endregion
#endif
            catch (Exception e)
            {
                str_exception += "file : " + filePath + "\r\n\r\n";
                str_exception += "error : " + e.ToString();
                // xlsxとdocxのみファイルを開いている状態でファイルアクセスした場合
                // ファイルが開かれている場合のエラー
                error_reason = ListForm.LIST_VIEW_NA;

                if (excel != null) excel.Close();
                if (word != null) word.Close();
                if (ppt != null) ppt.Close();

                return retTable;
            }
            return retTable;
        }

#if false
#region HyperLink修復

        /// <summary>
        /// OFFICEドキュメント内の破損しているハイパーリンクを置き換える文字列を返します
        /// </summary>
        /// <param name="brokenUri"></param>
        /// <returns></returns>
        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }
#endregion
#endif

        /// <summary>
        /// 取得したプロパティを画面クラスへセットします
        /// </summary>
        /// <param name="propertiesTable"></param>
        /// <param name="clsListData"></param>
        /// <param name="bStamp"></param>
        //public bool setProperties(Dictionary<string, string> propertiesTable, ref ListForm.GridListData clsListData, ref bool bStamp, string myOfficeCode)
        public bool setProperties(Dictionary<string, string> propertiesTable, ref ListForm.GridListData clsListData, ref bool bStamp, string myOfficeCode , string error_reason)  // 20171011 修正（開いているファイルを読込時に文書分類欄に「読取専用」表示）
        {
            bool bRet = false;

            try
            {
                // 20171009 コメントアウト（第3パラメータ「無」or「no_stamp」ファイル対応）
                // カテゴリーから読み込み
                // clsListData.strClassNo = (propertiesTable[PropertiesKeyList.STR_CATEGORY] != null) ? propertiesTable[PropertiesKeyList.STR_CATEGORY] : "";
                // 20180109 TSE kitada カテゴリキーが含まれていない場合はExceptionが出てしまうためチェックしてから取得に変更
                clsListData.strClassNo = propertiesTable.ContainsKey(PropertiesKeyList.STR_CATEGORY) ? propertiesTable[PropertiesKeyList.STR_CATEGORY] : "";

                // タグ情報があった場合は機密区分、文書分類番号、スタンプの有無をセットする
                if (propertiesTable.ContainsKey(PropertiesKeyList.STR_KEYWORDS))
                {
                    string[] strPropertyData = propertiesTable[PropertiesKeyList.STR_KEYWORDS].Split(';');

                    // 20171013 コメントアウト（保存期限、保存年数は属性設定.csvから取得する為、不要）
                    //// 機密区分、文書分類番号、スタンプの有無、保存期限、保存年数の5項目以上ある場合
                    //if (strPropertyData.Count() >= 5)
                    //{
                    //    clsListData.strSaveLives = strPropertyData[4].Trim();
                    //}
                    //// 機密区分、文書分類番号、スタンプの有無、保存期限の4項目以上ある場合
                    //if (strPropertyData.Count() >= 4)
                    //{
                    //    clsListData.strShelfLife = strPropertyData[3].Trim();
                    //}

                    // 機密区分、文書分類番号、スタンプの有無の3項目以上ある場合
                    if (strPropertyData.Count() >= 3)
                    {
                        if (strPropertyData[2].Trim() == ListForm.NO_STAMP) bStamp = false;
                    }
                    // 機密区分、文書分類番号の2項目以上ある場合
                    if (strPropertyData.Count() >= 2)
                    {
                        // 20170928 追加 TSE kitada
                        // 他事業所の場合は文書分類欄に他事業所名を表示する仕様へ変更。
                        // これに伴い、正式な文書分類はstrClassNo_hidden(グリッドの隠しセル)に格納し
                        // 従来の文書分類欄には文書分類または事業所コードを表示する用に修正

                        // 20171012 追加
                        // パラメータが2つで
                        if (strPropertyData.Count() == 2)
                        {
                            clsListData.strClassNo = strPropertyData[1].Trim();
                        }
                        // パラメータの3つ目が「no_stamp」or「無」で
                        else if (strPropertyData[2].Trim() == ListForm.NO_STAMP || strPropertyData[2].Trim() == "")
                        {
                            clsListData.strClassNo = strPropertyData[1].Trim();
                        }
                        // パラメータの3つ目が自事業所名でない場合
                        else if (strPropertyData.Length >= 3 && strPropertyData[2].Trim() != myOfficeCode)
                        {
                            clsListData.strClassNo = strPropertyData[2].Trim();
                        }
                        else
                        {
                            clsListData.strClassNo = strPropertyData[1].Trim();
                        }
                        clsListData.strClassNo_hideen = strPropertyData[1].Trim();
                    }
                    // 機密区分の1項目以上ある場合
                    if (strPropertyData.Count() >= 1)
                    {
                        clsListData.strSecrecyLevel = strPropertyData[0].Trim();
                    }
                }

                //if (propertiesTable[PropertiesKeyList.STR_CREATOR] != null)
                if (propertiesTable.ContainsKey(PropertiesKeyList.STR_CREATOR))
                {
                    // 作成者
                    clsListData.strCreator = propertiesTable[PropertiesKeyList.STR_CREATOR];
                }

                //if (propertiesTable[PropertiesKeyList.STR_LAST_MODIFIED_BY] != null)
                if (propertiesTable.ContainsKey(PropertiesKeyList.STR_LAST_MODIFIED_BY))
                {
                    // 最終更新者
                    clsListData.strLastModifiedBy = propertiesTable[PropertiesKeyList.STR_LAST_MODIFIED_BY];
                }

                bRet = true;                            
            }
            catch(Exception e)
            {
                // 20171228 追加（エラー理由保存）
                error_reason = ListForm.LIST_VIEW_NA;
            }

            // 20171011 修正（開いているファイルを読込時に文書分類欄に「読取専用」表示）
            if (error_reason != "")
            {
                clsListData.strClassNo = error_reason;
            }

            // 20171013 追加（文書分類検索用の隠しカラムに文書分類をそのままコピー）
            clsListData.strClassNoSerch_hideen = clsListData.strClassNo;

            return bRet;
        }
    }

#if false
#region HyperLink修復


    /// <summary>
    /// OFFICE ドキュメント内の破損したハイパーリンクを修復するためのクラス
    /// ref : http://ericwhite.com/blog/handling-invalid-hyperlinks-openxmlpackageexception-in-the-open-xml-sdk/
    /// </summary>
    public static class UriFixer
    {
        public static void FixInvalidUri(Stream fs, Func<string, Uri> invalidUriHandler)
        {
            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                foreach (var entry in za.Entries.ToList())
                {
                    if (!entry.Name.EndsWith(".rels"))
                        continue;
                    bool replaceEntry = false;
                    XDocument entryXDoc = null;
                    using (var entryStream = entry.Open())
                    {
                        try
                        {
                            entryXDoc = XDocument.Load(entryStream);
                            if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                            {
                                var urisToCheck = entryXDoc
                                    .Descendants(relNs + "Relationship")
                                    .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                                foreach (var rel in urisToCheck)
                                {
                                    var target = (string)rel.Attribute("Target");
                                    if (target != null)
                                    {
                                        try
                                        {
                                            Uri uri = new Uri(target);
                                        }
                                        catch (UriFormatException)
                                        {
                                            Uri newUri = invalidUriHandler(target);
                                            rel.Attribute("Target").Value = newUri.ToString();
                                            replaceEntry = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch (XmlException)
                        {
                            continue;
                        }
                    }
                    if (replaceEntry)
                    {
                        var fullName = entry.FullName;
                        entry.Delete();
                        var newEntry = za.CreateEntry(fullName);
                        using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                        using (XmlWriter xmlWriter = XmlWriter.Create(writer))
                        {
                            entryXDoc.WriteTo(xmlWriter);
                        }
                    }
                }
            }
        }
    }
#endregion
#endif

}
