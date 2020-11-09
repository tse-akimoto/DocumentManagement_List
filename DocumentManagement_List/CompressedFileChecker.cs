using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Ionic.Zip;
using System.Windows.Forms;
using Microsoft.Office.Core;
using DocumentManagement_List.Properties;   // step2 iwasa

namespace DocumentManagement_List
{
    class CompressedFileChecker
    {
        #region <内部変数>

        /// <summary>
        /// 共通設定項目
        /// </summary>
        public SettingForm settingForm { set; get; }

        /// <summary>
        /// zipファイルパスリスト
        /// </summary>
        public Dictionary<string, HashSet<string>> dicZipPath = new Dictionary<string, HashSet<string>>();

        /// <summary>
        /// 拡張子リスト
        /// </summary>
        public static List<string> listExtension = new List<string>();  // step2 iwasa

        #endregion

        #region <クラス定義>

        /// <summary>
        /// 選択したzipファイルの解凍先のリストを取得
        /// </summary>
        /// <param name="fileName">ファイル名</param>
        /// <param name="dicPasswordZip">パスワード付きzipのリスト</param>
        /// <param name="dicErrorZip">エラーzipのリスト</param>
        /// <returns>解凍に成功したzipのリスト</returns>
        public List<string> GetUnzipFilePathList(string fileName,
            ref Dictionary<string, List<string>> dicPasswordZip,
            ref Dictionary<string, List<string>> dicErrorZip
            )
        {
            List<string> result = new List<string>();

            var options = new ReadOptions
            {
                StatusMessageWriter = System.Console.Out,

                // 多言語対応のため、変更
                Encoding = System.Text.Encoding.Default
            };

            try
            {
                string NonDrivePath = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName));

                // ドライブ除外
                NonDrivePath = NonDrivePath.Remove(0, 3);

                string TempPath = Path.Combine(settingForm.clsCommonSettting.strTempPath, NonDrivePath);


                if (!Directory.Exists(TempPath))
                {
                    Directory.CreateDirectory(TempPath);
                }

                using (ZipFile Zip = ZipFile.Read(fileName, options))
                {

                    var Split = Zip[0].FileName.Split('/');
                    string ZipFolderPath = Path.Combine(TempPath, Split[0]);

                    if (Split.Count() == 1)
                    {
                        // Zip内フォルダが存在しない場合
                        TempPath += @"\" + Path.GetFileNameWithoutExtension(fileName);
                        ZipFolderPath = TempPath;
                    }

                    result.Add(ZipFolderPath);

                    // パスワード付きzipファイルの中身の情報を取得する step2 iwasa
                    GetPasswordZipFileList(fileName, Zip.Info, ref dicPasswordZip);

                    // 展開先に同名のファイルがあれば上書きする
                    Zip.ExtractExistingFile = Ionic.Zip.ExtractExistingFileAction.OverwriteSilently;
                    // ZIPファイル内の全てのファイルを解凍する
                    Zip.ExtractAll(TempPath);

                    foreach (var entry in Zip.Entries)
                    {
                        string outputPath = Path.GetFullPath(Path.Combine(TempPath, entry.FileName));
                        result.Add(outputPath);
                    }
                }

                dicZipPath[fileName] = new HashSet<string>(result);
            }
            catch (Ionic.Zip.BadPasswordException ex)
            {
                // なにもしない
            }
            catch (Exception _ex)
            {
                // パスワード以外の例外
                // 解凍不可に追加
                dicErrorZip.Add(fileName, result);
            }

            return result;
        }

        /// <summary>
        /// 対象がzipか
        /// </summary>
        /// <param name="s">文字列</param>
        /// <returns>true:zip以外 false:zip</returns>
        public static bool judge(string s)
        {
            return !s.Contains("zip");
        }

        /// <summary>
        /// Zip内パスをすべて取得する
        /// </summary>
        /// <param name="listCopyBuf">ファイルパスリスト</param>
        /// <param name="dicZipResult">ZIPリスト</param>
        /// <param name="dicCompressedItem">全てのパス</param>
        /// <param name="dicPasswordZip">パスワード付きZIP</param>
        /// <param name="dicErrorZip">エラーZIP</param>
        public void GetZipAllList(
            List<string> listCopyBuf,
            ref Dictionary<string, Dictionary<string, HashSet<string>>> dicZipResult,
            ref Dictionary<string, HashSet<string>> dicCompressedItem,
            ref Dictionary<string, List<string>> dicPasswordZip,
            ref Dictionary<string, List<string>> dicErrorZip
            )
        {
            foreach (string BufFile in listCopyBuf)
            {
                HashSet<string> listCompressed = new HashSet<string>();

                if (Path.GetExtension(BufFile).Contains("zip") != false)
                {
                    // 解凍してリストに追加
                    listCompressed = new HashSet<string>(GetUnzipFilePathList(BufFile, ref dicPasswordZip, ref dicErrorZip));

                    // zipリスト
                    dicZipResult[BufFile] = new Dictionary<string, HashSet<string>>(dicZipPath);
                    dicZipPath.Clear();

                    // すべてのパス
                    dicCompressedItem[BufFile] = listCompressed;
                }
            }
        }

        /// <summary>
        /// 選択したzipファイルを再zip化する
        /// </summary>
        /// <param name="listZipTarget">ZIP対象</param>
        /// <param name="dicZipResult">上書き対象のZIPリスト</param>
        public void SelectZipProc(
            ref HashSet<string> listZipTarget,
            ref Dictionary<string, Dictionary<string, HashSet<string>>> dicZipResult
            )
        {
            foreach (string ZipTarget in listZipTarget)
            {
                // 上書き対象のZIPファイル
                var InputList = dicZipResult[ZipTarget];

                foreach (var list in InputList)
                {
                    ZipFile _zip = new ZipFile(System.Text.Encoding.Default);
                    string ZipFilePath = list.Key.ToString();
                    List<string> listItem = list.Value.ToList();
                    string ZipDir = listItem[0];

                    // 圧縮レベルの設定
                    _zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;
                    // 必要な時はZIP64で圧縮する
                    _zip.UseZip64WhenSaving = Zip64Option.AsNecessary;

                    // Zip化
                    _zip.AddDirectory(ZipDir);
                    _zip.Save(ZipFilePath);
                }
            }
        }

        /// <summary>
        /// tmpフォルダを削除する
        /// </summary>
        /// <param name="formSetting">共通設定項目</param>
        /// <param name="isFormClose">フォームの開閉状態</param>
        static public void ResetTempFolder(
            SettingForm formSetting,
            bool isFormClose
            )
        {
            if (string.IsNullOrEmpty(formSetting.clsCommonSettting.strTempPath) == false)
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(formSetting.clsCommonSettting.strTempPath);
                if (di.Exists != false)
                {
                    di.Delete(true);
                }

                // フォームクローズではないとき
                if (isFormClose == false)
                {
                    di.Create();
                }
            }
        }

        /// <summary>
        /// zip内ファイルの存在をチェック
        /// </summary>
        /// <param name="dic">ファイルパスリスト</param>
        /// <param name="CheckPath">zip内ファイルパス</param>
        /// <param name="TargetZipPath">処理結果パス</param>
        static public bool IsCompresstionZipItem(
            Dictionary<string, HashSet<string>> dic,
            string CheckPath,
            out string TargetZipPath)
        {
            bool ret = false;
            string ZipPath = "";

            // Zip内のデータか？
            foreach (var list in dic)
            {
                foreach (string zipItem in list.Value)
                {
                    try
                    {
                        if (File.GetAttributes(zipItem).HasFlag(FileAttributes.Directory) == false)
                        {
                            if (CheckPath.Contains(zipItem))
                            {
                                ret = true;
                                ZipPath = list.Key;
                            }
                        }
                    }
                    catch
                    {
                        // ファイル・フォルダではない
                    }
                }
            }

            TargetZipPath = ZipPath;

            return ret;
        }

        /// <summary>
        /// パスワード付きzip内ファイルをチェック
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="info"></param>
        /// <param name="dicPasswordZip"></param>
        /// <returns>true: パスワード付き false:パスワードなし</returns>
        private bool GetPasswordZipFileList(string fileName, string info, ref Dictionary<string, List<string>> dicPasswordZip)    // step2 iwasa
        {
            bool _isPassword = false;
            List<string> lstFileName = new List<string>();

            // 複数行の文字列を解析するため、一行ごとに分ける
            string[] lines = info.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            foreach (string item in lines)
            {
                if (item.Contains("ZipEntry") == true)
                {
                    if (listExtension.Contains(Path.GetExtension(item)) == true)
                    {
                        // 対象の拡張子であればファイル名を追加する
                        lstFileName.Add(Path.GetFileName(item));
                    }
                }
                else if (item.Contains("Encrypted") == true)
                {
                    if (item.Contains("True") == true)
                    {
                        // パスワード付きzipのため解凍不可
                        _isPassword = true;
                    }
                }
            }

            if (_isPassword)
            {
                // 解凍不可zip
                // fileNameには解凍不可zipのフルパスが入る
                // lisFileNameにはzip内のファイル名が入る
                dicPasswordZip.Add(fileName, lstFileName);
            }

            return _isPassword;
        }

        #endregion
    }
}
