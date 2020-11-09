using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using DSOFile;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Globalization; // step2    iwasa
using System.Threading;
using DocumentManagement_List.Properties;
using System.Collections.ObjectModel;

namespace DocumentManagement_List
{
    public partial class ListForm : Form
    {
        #region <定数定義>

        /// <summary>
        /// 検索警告ファイル数
        /// </summary>
        public const int SEARCH_ALERT_COUNT = 1000;

        /// <summary>
        /// 読み込み可能なファイル数上限 10万
        /// </summary>
        /// <remarks>20190705 TSE matsuo 追加</remarks>
        private const int READ_FILE_COUNT_LIMIT = 100000;

        /// <summary>
        /// ファイル無し
        /// </summary>
        public const int EXTENSION_NONE = 0;

        /// <summary>
        /// PDFファイル
        /// </summary>
        public const int EXTENSION_PDF = 1;

        /// <summary>
        /// EXCELファイル
        /// </summary>
        public const int EXTENSION_EXCEL = 2;

        /// <summary>
        /// WORDファイル
        /// </summary>
        public const int EXTENSION_WORD = 3;

        /// <summary>
        /// POWERPOINTファイル
        /// </summary>
        public const int EXTENSION_POWERPOINT = 4;

        /// <summary>
        /// zipファイル
        /// </summary>
        public const int EXTENSION_ZIP = 5;

        /// <summary>
        /// タイプなし
        /// </summary>
        public const int EXTENSION_TYPE_NONE = 0;

        /// <summary>
        /// OFFICEファイル
        /// </summary>
        public const int EXTENSION_TYPE_OFFICE = 1;

        /// <summary>
        /// PDFファイル
        /// </summary>
        public const int EXTENSION_TYPE_PDF = 2;

        /// <summary>
        /// OFFICE+PDFファイル
        /// </summary>
        public const int EXTENSION_TYPE_OFFICEPDF = 3;

        /// <summary>
        /// 全て
        /// </summary>
        public const int EXTENSION_TYPE_ALL = 4;    // step2 isasa

        // ファイルプロパティ関連

        /// <summary>
        /// プロパティに書き込むSAB秘 S秘
        /// </summary>
        public const string SECRECY_PROPERTY_S = "SecrecyS";

        /// <summary>
        /// プロパティに書き込むSAB秘 A秘
        /// </summary>
        public const string SECRECY_PROPERTY_A = "SecrecyA";

        /// <summary>
        /// プロパティに書き込むSAB秘 B秘
        /// </summary>
        public const string SECRECY_PROPERTY_B = "SecrecyB";

        /// <summary>
        /// プロパティに書き込むスタンプ無し情報
        /// </summary>
        public const string NO_STAMP = "no_stamp";

        /// <summary>
        /// 区分区切り文字
        /// </summary>
        public const string SEPARATE = ";";

        // リスト関連

        /// <summary>
        /// リスト表示列数
        /// </summary>
        public const int MAX_LIST_VIEW_COLUMNS = 8;

        /// <summary>
        /// 文書分類表示：該当無し
        /// </summary>
        public const string LIST_VIEW_NA = "N/A";


        // ユーザー設定関連

        /// <summary>
        /// ユーザー設定格納フォルダ名（新）
        /// </summary>
        public const string USER_SETFOLDERNAME = "SAB";

        /// <summary>
        /// ファイル名
        /// </summary>
        public const string USER_SETFILENAME = "user_setting_list.config";

        /// <summary>
        /// ユーザー設定格納フォルダ名（旧）
        /// </summary>
        public const string USER_SET_OLDFOLDERNAME = @"Microsoft Corporation\Microsoft Office 2010";  // 20171009 追加 ユーザー設定格納フォルダ名（旧）


        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN1 = 1;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN2 = 2;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN3 = 3;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN4 = 4;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN5 = 5;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN6 = 6;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN7 = 7;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN8 = 8;    // 20170905 追加（保存期限、保存年数の追加）

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN9 = 9;    // 20170905 追加（保存期限、保存年数の追加）

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN10 = 10;    // 20170907 追加（作成者、最終更新者の追加）

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_SHEET_COLUMN11 = 11;    // 20170907 追加（作成者、最終更新者の追加）

        /// <summary>
        /// 属性設定CSVの分類コードの列位置
        /// </summary>
        public const int ZOKUSEI_CSV_BUNRUI_INDEX = 5;

        /// <summary>
        /// 属性設定CSVの保持期限の列位置
        /// </summary>
        public const int ZOKUSEI_CSV_KIGEN_INDEX = 6;

        /// <summary>
        /// 属性設定CSVの日数の列位置
        /// </summary>
        public const int ZOKUSEI_CSV_NISSU_INDEX = 7;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_READDATA_COUNT = 10;

        /// <summary>
        /// Excel入出力関連
        /// </summary>
        public const int EXCEL_OUTPUT_COLUMN_COUNT = 9;

        /// <summary>
        /// PDFテンポラリファイルパスフォーマット
        /// </summary>
        public const string PDF_TEMPFILEPATH_FORMAT = "{0}DML_PDF_TEMP{1}.PDF";

        /// <summary>
        /// 除外ファイルログ出力先
        /// </summary>
        public const string EXCLUSION_LOG_PATH = @"C:\ProgramData\SAB";

        /// <summary>
        /// 共通設定読込
        /// </summary>
        SettingForm settingForm;

        /// <summary>
        /// 共通設定ファイルに不足項目がないかチェック
        /// </summary>
        bool CommonSetttingFlg = true;

        /// <summary>
        /// 共通設定ファイルで不足している項目を表示するメッセージ
        /// </summary>
        string CommonSetttingMessage = null;

        /// <summary>
        /// 検索フォルダ内（サブフォルダ含め全て）にシステムファイルがあった場合：true
        /// </summary>
        public bool accessSystemFileFlg = false;

        /// <summary>
        /// 表示中のファイルリスト
        /// </summary>
        private ArrayList CurrentArrayGridOutput = new ArrayList();

        /// <summary>
        /// 表示中のZIPリスト数
        /// </summary>
        private int CurrentViewZipCount = 0;

        /// <summary>
        /// ソート状態
        /// </summary>
        private static ListSortDirection ZipSortDirection = ListSortDirection.Descending;

        /// <summary>
        /// ソート状態
        /// </summary>
        private static ListSortDirection FileNameSortDirection = ListSortDirection.Descending;

        /// <summary>
        /// ソート状態
        /// </summary>
        private static ListSortDirection CreateTimeSortDirection = ListSortDirection.Descending;

        /// <summary>
        /// ソート状態
        /// </summary>
        private static ListSortDirection UpdateTimeSortDirection = ListSortDirection.Descending;

        /// <summary>
        /// ソート状態
        /// </summary>
        private static ListSortDirection ClassNoSortDirection = ListSortDirection.Descending;

        /// <summary>
        /// ソート状態
        /// </summary>
        private static ListSortDirection SecrecyLevelSortDirection = ListSortDirection.Descending;

        /// <summary>
        /// ソート状態
        /// </summary>
        private static ListSortDirection FilePathSortDirection = ListSortDirection.Descending;

        /// <summary>
        /// データグリッドビュー列インデックス列挙体
        /// </summary>
        private enum DataGridViewColumnIndex
        {
            /// <summary>
            /// zip形式
            /// </summary>
            ZipFormat = 0,  // step2 iwasa

            /// <summary>
            /// ファイル名
            /// </summary>
            FileName,

            /// <summary>
            /// ファイル種別
            /// </summary>
            FileType,

            /// <summary>
            /// 作成日
            /// </summary>
            CreatedDate,

            /// <summary>
            /// 更新日
            /// </summary>
            UpdatedDate,

            /// <summary>
            /// 選択(チェックボックス)
            /// </summary>
            Select,

            /// <summary>
            /// 文書分類
            /// </summary>
            DocumentType,

            /// <summary>
            /// 機密区分
            /// </summary>
            SecretType,

            /// <summary>
            /// ファイルパス
            /// </summary>
            FilePath,

            /// <summary>
            /// 
            /// </summary>
            Blank1,

            /// <summary>
            /// 
            /// </summary>
            Blank2,

            /// <summary>
            /// 作成者
            /// </summary>
            CreatedBy,

            /// <summary>
            /// 最終更新者
            /// </summary>
            UpdatedBy,

            /// <summary>
            /// 機密区分(非表示)
            /// </summary>
            SecretTypeHidden,

            /// <summary>
            /// Zipファイルtemp先(非表示)
            /// </summary>
            ZipFilePath,

            /// <summary>
            /// Zipファイル個数(非表示)
            /// </summary>
            ZipFileCount
        }

        #endregion

        #region <内部変数>

        /// <summary>
        /// 検索対象拡張子
        /// </summary>
        public string[] StrExtension_Narrow;

        /// <summary>
        /// OpenXML対象の拡張子
        /// </summary>
        public string[] StrOpenXML_Narrow;

        /// <summary>
        /// 拡張子別種類
        /// </summary>
        public int[] iExtension_Narrow;

        /// <summary>
        /// 中止フラグ
        /// </summary>
        private Boolean bStopFlg;

        /// <summary>
        /// 割り込み禁止フラグ
        /// </summary>
        private Boolean bInterruptDisabled;

        /// <summary>
        /// リストのファイルタイプ
        /// </summary>
        private int IListFileType;

        /// <summary>
        /// リストのハッシュテーブル(同じファイルパス対策)
        /// </summary>
        private Hashtable htList;

        /// <summary>
        /// 内部保持用ファイルリスト
        /// </summary>
        ArrayList ArrayFileList = new ArrayList();

        // 20200804 追加 ZIP解凍関連

        /// <summary>
        /// ZIPファイルリスト
        /// </summary>
        Dictionary<string, Dictionary<string, HashSet<string>>> dicZipResult = new Dictionary<string, Dictionary<string, HashSet<string>>>();

        /// <summary>
        /// 全てのパス
        /// </summary>
        Dictionary<string, HashSet<string>> dicCompressedItem = new Dictionary<string, HashSet<string>>();

        /// <summary>
        /// パスワード付きZIPリスト
        /// </summary>
        Dictionary<string, List<string>> dicPasswordZip = new Dictionary<string, List<string>>();

        /// <summary>
        /// エラーZIPリスト
        /// </summary>
        Dictionary<string, List<string>> dicErrorZip = new Dictionary<string, List<string>>();

        #endregion

        #region <クラス定義>

        /// <summary>
        /// リストデータ設定
        /// </summary>
        public class GridListData
        {
            /// <summary>
            /// zip形式
            /// </summary>
            public string strZipFormat;

            /// <summary>
            /// ファイル名
            /// </summary>
            public string strFileName;

            /// <summary>
            /// 種類
            /// </summary>
            public string strFileType;

            /// <summary>
            /// 作成日
            /// </summary>
            public string strCreateDate;

            /// <summary>
            /// 更新日
            /// </summary>
            public string strUpdateDate;

            /// <summary>
            /// 文書分類
            /// </summary>
            public string strClassNo;

            /// <summary>
            /// 機密区分
            /// </summary>
            public string strSecrecyLevel;

            /// <summary>
            /// ファイルパス
            /// </summary>
            public string strFilePath;

            /// <summary>
            /// 作成者
            /// </summary>
            public string strCreator;

            /// <summary>
            /// 最終更新者
            /// </summary>
            public string strLastModifiedBy;

            /// <summary>
            /// 文書分類(隠し)
            /// </summary>
            public string strClassNo_hideen;

            /// <summary>
            /// 文書分類検索用(隠し)
            /// </summary>
            public string strClassNoSerch_hideen;

            /// <summary>
            /// ZIPファイルtmp先(隠し)
            /// </summary>
            public string strTmpZipFilePath;


            /// <summary>
            /// コンストラクタ
            /// </summary>
            public GridListData()
            {
                // 初期化
                strZipFormat = "";
                strFileName = "";
                strFileName = "";
                strFileType = "";
                strCreateDate = "";
                strUpdateDate = "";
                strClassNo = "";
                strSecrecyLevel = "";
                strFilePath = "";
                strCreator = "";        // 20170907 追加 (作成者)
                strLastModifiedBy = ""; // 20170907 追加 (最終更新者)
                strClassNo_hideen = ""; // 20170928 追加 (文書分類(隠し))
                strClassNoSerch_hideen = ""; // 20171012 追加 (事業所名(隠し))
                strTmpZipFilePath = "";        // 20200805 追加 (ZIPファイルtmp先(隠し))
            }
        }
        #endregion

        #region <コンストラクタ>

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ListForm()
        {
            // 共通設定ファイル読み込み
            settingForm = new SettingForm();

            if (settingForm.clsCommonSettting == null)
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += Resources.msgCommonFileNothing + Environment.NewLine;
            }
            else
            {
                CommonSetttingMessage = Resources.msgCommonSettingNothing + Environment.NewLine;
                // 共通設定ファイルに不足項目がないかチェック
                if (string.IsNullOrEmpty(settingForm.clsCommonSettting.strDefaultSecrecyLevel))   // デフォルト機密区分
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonDefaultSecrecy + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(settingForm.clsCommonSettting.strOfficeCode))   // 事業所コード
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonOfficeCode + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(settingForm.clsCommonSettting.strCulture))   // 言語設定
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonCulture + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(settingForm.clsCommonSettting.strSABListLocalPath))   // 文書のローカルパス
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonSABListLocalPath + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(settingForm.clsCommonSettting.strSABListServerPath))   // 文書のサーバーパス
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonSABListServerPath + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(settingForm.clsCommonSettting.strTempPath))   // zip一時解凍先
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonTempPath + Environment.NewLine;
                }
                if (settingForm.clsCommonSettting.lstSecureFolder.Count == 0)   // セキュアフォルダリスト
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonSecureFolder + Environment.NewLine;
                }
                if (settingForm.clsCommonSettting.lstFinal.Count == 0)   // 「最終版」を表す文字列
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Resources.msgCommonFinal + Environment.NewLine;
                }
            }

            // 共通設定ファイルに不足項目がある場合、何もしない
            if (CommonSetttingFlg)
            {
                // 言語設定読込み  // step2    iwasa
                CultureInfo culture = CultureInfo.GetCultureInfo(settingForm.clsCommonSettting.strCulture);
                Thread.CurrentThread.CurrentUICulture = culture;

                // 検索対象ファイル設定
                StrExtension_Narrow = new string[] {
                ".pdf",
                ".xlsx", ".xlsm", ".xls",
                ".docx", ".doc",
                ".pptx", ".ppt"
            };
                iExtension_Narrow = new int[] {
                EXTENSION_PDF,
                EXTENSION_EXCEL, EXTENSION_EXCEL, EXTENSION_EXCEL,
                EXTENSION_WORD, EXTENSION_WORD,
                EXTENSION_POWERPOINT, EXTENSION_POWERPOINT
            };

                // OpenXMLでの読み書き対象
                StrOpenXML_Narrow = new string[]
                {
                ".xlsx", ".xlsm", ".docx", ".pptx"
                };

                // グローバル変数初期化
                bStopFlg = false;                                       // 中止フラグOFF
                bInterruptDisabled = false;                             // 割り込み禁止フラグOFF

                // ドキュメントの更新が必要なら更新する
                settingForm.UpdateDocument();

                // ハッシュテーブル宣言
                htList = new Hashtable();

                CompressedFileChecker.listExtension = StrExtension_Narrow.ToList(); // step2 iwasa
            }

            InitializeComponent();
        }
        #endregion

        #region <フォームイベント>

        /// <summary>
        /// フォームロード
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListForm_Load(object sender, EventArgs e)
        {
            if (!CommonSetttingFlg)
            {
                // 共通設定ファイルに不足項目あり
                MessageBox.Show(CommonSetttingMessage,
                    Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);

                // 強制終了
                Environment.Exit(0x8020);
            }

            // 画面初期化
            Initialize();
        }

        /// <summary>
        /// フォームクローズ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            // 20200806 tempファイル削除
            CompressedFileChecker.ResetTempFolder(settingForm, true);
        }

        /// <summary>
        /// 参照ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonReference_Click(object sender, EventArgs e)
        {
            var dialog = new CommonOpenFileDialog();

            // ダイアログのタイトル
            dialog.Title = Resources.msgSelectSearchFolder; // step2    iwasa

            // 初期ディレクトリに指定
            dialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // フォルダ指定がありそのフォルダが存在している場合はそのパスを初期ディレクトリに設定する
            if ((textBoxFolderPath.Text != "") && (System.IO.Directory.Exists(textBoxFolderPath.Text) == true))
            {
                dialog.InitialDirectory = textBoxFolderPath.Text;
            }

            // フォルダー指定設定
            dialog.IsFolderPicker = true;

            // Openダイアログを表示
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                textBoxFolderPath.Text = dialog.FileName;
            }
        }

        /// <summary>
        /// 更新年月日チェック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxUpdate_CheckedChanged(object sender, EventArgs e)
        {
            // 更新年月日の入力許可を設定
            dateTimePickerUpdateFrom.Enabled = checkBoxUpdate.Checked;
            dateTimePickerUpdateTo.Enabled = checkBoxUpdate.Checked;
        }

        /// <summary>
        /// 作成年月日チェック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxCreate_CheckedChanged(object sender, EventArgs e)
        {
            // 作成年月日の入力許可を設定
            dateTimePickerCreateFrom.Enabled = checkBoxCreate.Checked;
            dateTimePickerCreateTo.Enabled = checkBoxCreate.Checked;
        }

        /// <summary>
        /// 全てチェックボックス変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxSAB_All_CheckedChanged(object sender, EventArgs e)
        {
            // 割り込み禁止でない場合
            if (!bInterruptDisabled)
            {
                // 全てがチェックされた場合は他のチェックボックスも同じくチェックを変更する(割り込み禁止使用)
                bInterruptDisabled = true;
                checkBoxSAB_S.Checked = checkBoxSAB_All.Checked;
                checkBoxSAB_A.Checked = checkBoxSAB_All.Checked;
                checkBoxSAB_B.Checked = checkBoxSAB_All.Checked;
                checkBoxSAB_None.Checked = checkBoxSAB_All.Checked;
                checkBoxSAB_Other.Checked = checkBoxSAB_All.Checked;
                bInterruptDisabled = false;
            }
        }

        /// <summary>
        /// SABチェックボックス変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxSAB_CheckedChanged(object sender, EventArgs e)
        {
            // 割り込み禁止でない場合
            if (!bInterruptDisabled)
            {
                // 全ての行にチェックがついている場合は全選択チェックボックスをチェックする。そうでない場合はチェックを外す。(割り込み禁止使用)
                bInterruptDisabled = true;
                // 20200811 STEP2
                checkBoxSAB_All.Checked = ((checkBoxSAB_S.Checked) && (checkBoxSAB_A.Checked) && (checkBoxSAB_B.Checked) && (checkBoxSAB_Other.Checked) && (checkBoxSAB_None.Checked)) ? true : false;
                bInterruptDisabled = false;
            }
        }

        /// <summary>
        /// 検索ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSearch_Click(object sender, EventArgs e)
        {
            string[] stFileList;                                    // ファイルリスト
            bool fileReadErrFlg = false;                          // ファイル読込時にエラーが発生した場合:true

            // 20170801 追加（システムファイルへのアクセスフラグをリセット）
            accessSystemFileFlg = false;

            #region<ファイル検索処理>
            try
            {
                #region<入力チェック>
                // フォルダが入力されているか確認
                if (textBoxFolderPath.Text == "")
                {
                    MessageBox.Show(Resources.msgSelectFolder, Resources.msgConfirmation, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    textBoxFolderPath.Focus();
                    return;
                }

                // フォルダ (ディレクトリ) が存在しているかどうか確認する
                if (!System.IO.Directory.Exists(textBoxFolderPath.Text))
                {
                    MessageBox.Show(Resources.msgNotExistFolder, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    textBoxFolderPath.Focus();
                    return;
                }

                // 絞込み条件デフォルトチェック
                if ((radioButtonTypeAll.Checked == false) ||    // step2 iwasa
                     (textBoxFileName.Text != "") ||
                     (checkBoxCreate.Checked == true) ||
                     (checkBoxUpdate.Checked == true) ||
                     (checkBoxSAB_All.Checked == false) ||
                     (textBoxClassNo.Text != "")
                    )
                {

                    DialogResult dr = MessageBox.Show(
                            Resources.msgCheckClearCondition,
                            Resources.msgConfirmation,
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Exclamation
                            );

                    if (dr == System.Windows.Forms.DialogResult.No)  // step2    iwasa
                    {
                        return;
                    }
                }

                // 20170913追加（検索に時間が掛かる可能性のあるフォルダが指定されていた場合の処理）
                // 検索に時間が掛かるフラグ
                bool is_serch_long_time = false;
                // フォルダのパスに「\」が何個あるか確認
                int enMark_count = textBoxFolderPath.Text.Length - textBoxFolderPath.Text.Replace(@"\", "").Length;
                
                switch (enMark_count)
                {
                    // フォルダのパスに「\」が1つしかない場合
                    case 1: is_serch_long_time = true; break;
                    // フォルダのパスに「\」が3つあり、フォルダのパスの先頭に「\\」が2つある場合
                    case 3: if (textBoxFolderPath.Text.IndexOf(@"\\") == 0) is_serch_long_time = true; break;
                            default: break;
                }

                // 20170913追加（検索に時間が掛かる可能性のあるフォルダが指定されていた場合の処理）
                // 検索に時間が掛かると判定された場合
                // 20190705 TSE matsuo サブフォルダを含む場合に限定
                if (is_serch_long_time && radioButtonSubFolderInc.Checked)
                {
                    DialogResult dr = MessageBox.Show(
                        Resources.msgSelectFolderUpper + Environment.NewLine + Resources.msgSearchLongTime,
                        Resources.msgConfirmation,
                        MessageBoxButtons.YesNo,
                         MessageBoxIcon.Exclamation
                         );

                    if (dr == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }

                // 絞り込み条件初期化
                // All指定
                //radioButtonTypeOfficePdf.Checked = true;  step2 iwasa
                radioButtonTypeAll.Checked = true;

                // ファイル名指定
                textBoxFileName.Text = "";

                // 作成年月日
                checkBoxCreate.Checked = false;
                dateTimePickerCreateFrom.Value = DateTime.Today.AddYears(-1);
                dateTimePickerCreateTo.Value = DateTime.Today;

                // 更新年月日
                checkBoxUpdate.Checked = false;
                dateTimePickerUpdateFrom.Value = DateTime.Today.AddYears(-1);
                dateTimePickerUpdateTo.Value = DateTime.Today;

                // S/A/B秘指定
                checkBoxSAB_All.Checked = true;
                checkBoxSAB_S.Checked = true;
                checkBoxSAB_A.Checked = true;
                checkBoxSAB_B.Checked = true;
                checkBoxSAB_None.Checked = true;
                checkBoxSAB_Other.Checked = true;

                // 文書分類指定
                textBoxClassNo.Text = "";

                #endregion

                // 画面入力不可
                ProcControlEnabled(false);

                // 進捗表示
                //labelStatus.Text = "ファイル検索中 対象ファイル 0件";
                labelStatus.Text = string.Format(Resources.msgStatusSearch, 0);   // step2    iwasa

                System.Windows.Forms.Application.DoEvents();

                // ファイルリスト取得
                List<string> listFiles = new List<string>();
                //Boolean bConfirm = false; // step2 iwasa
                System.IO.SearchOption searchOption = (radioButtonSubFolderInc.Checked) ? System.IO.SearchOption.AllDirectories : System.IO.SearchOption.TopDirectoryOnly;


                // 20170911 修正（システムファイルの読込時にシステムファイル以外の検索結果は表示するよう修正）
                string error_dt = DateTime.Now.ToString("yyyyMMddHHmmss");  // エラーファイルに使用する日付
                List<string> error_log = new List<string>();
                List<string> listBufFiles = GetNotSystemFileList(textBoxFolderPath.Text, error_dt, ref error_log , radioButtonSubFolderInc.Checked);
                List<string> listCopyBuf = new List<string>(listBufFiles);

                // step2 iwasa
                // zip解凍前にファイル件数をカウントする
                // zipファイルは1ファイルとして扱う
                // 指定件数を超えた場合は続行の確認ダイアログを表示する
                bool IsZipTarget = checkBoxZipTarget.Checked;
                int numFiles = 0;
                List<string> listExtension = new List<string>();
                listExtension = StrExtension_Narrow.ToList();
                
                if(IsZipTarget != false)
                {
                    listExtension.Add(".zip");
                }

                foreach (string strFilePath in listBufFiles)
                {
                    // 拡張子の数だけループ
                    foreach (string exp in listExtension)
                    {
                        // 指定の拡張子以外の場合は登録しない（*.xlsで検索すると.xlsx、.xlsmが検索されてしまう対策）
                        // 一時ファイルの場合は登録しない
                        if ((Path.GetExtension(strFilePath) == exp) &&
                            (Path.GetFileName(strFilePath).Substring(0, 2) != "~$"))
                        {
                            numFiles++;
                        }
                    }
                }

                if (numFiles > SEARCH_ALERT_COUNT)
                {
                    // ファイル検索中 対象ファイル {0}件
                    labelStatus.Text = string.Format(Resources.msgStatusSearch, numFiles);

                    DialogResult dr = MessageBox.Show(
                        Resources.msgFileTarget + SEARCH_ALERT_COUNT + Resources.msgFileContinueProcess,
                        Resources.msgConfirmation,
                        MessageBoxButtons.YesNo,
                         MessageBoxIcon.Exclamation);

                    if (dr == System.Windows.Forms.DialogResult.No)
                    {
                        // 画面入力可
                        ProcControlEnabled(true);
                        return;
                    }
                }
                
                CompressedFileChecker.ResetTempFolder(settingForm, false);
                CompressedFileChecker zipTest = new CompressedFileChecker();
                zipTest.settingForm = settingForm;

                dicZipResult.Clear();
                dicCompressedItem.Clear();
                dicPasswordZip.Clear();     // step2 iwasa
                dicErrorZip.Clear();

                if (IsZipTarget != false)
                {
                    zipTest.GetZipAllList(listCopyBuf, ref dicZipResult, ref dicCompressedItem, ref dicPasswordZip, ref dicErrorZip);

                    foreach (var list in dicCompressedItem)
                    {
                        foreach (string path in list.Value)
                        {
                            listBufFiles.Add(path);
                        }
                    }
                }

                if (error_log.Count != 0)
                {
                    // カレントディレクトに「Log」フォルダが無ければ作成
                    if (!Directory.Exists(EXCLUSION_LOG_PATH + @"\Log"))
                    {
                        Directory.CreateDirectory(EXCLUSION_LOG_PATH + @"\Log");
                    }

                    // エラーログファイル名「errorLog_yyyymmddHHMMss.log」
                    using (StreamWriter writer = new StreamWriter(EXCLUSION_LOG_PATH + @"\Log" + @"\errorLog_" + error_dt + ".log", true))    // 末尾に追記
                    {
                        // エラーあり、ファイル出力
                        foreach (string err in error_log)
                        {
                            // エラーログファイルに書込（末尾に追記）
                            writer.WriteLine(err);
                        }
                        writer.Close();
                    }
                }

                // 2019705 TSE matsuo ファイル数が10万件を超えている場合エラーとする
                if(listBufFiles.Count > READ_FILE_COUNT_LIMIT)
                {
                    MessageBox.Show(
                        Resources.msgFileOverHundredThousand + Environment.NewLine
                        + Resources.msgFileLessHundredThousand
                        , Resources.msgError
                        , MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    // 画面入力可
                    ProcControlEnabled(true);
                    return;
                }

                // ファイルパスの数だけループ
                foreach (string strFilePath in listBufFiles)
                {
                    // 待機
                    Delay(0.01);

                    // 中止ボタン押下した場合は処理終了
                    if (bStopFlg == true) break;

                    // 拡張子の数だけループ
                    foreach (string exp in listExtension)
                    {
                        // 指定の拡張子以外の場合は登録しない（*.xlsで検索すると.xlsx、.xlsmが検索されてしまう対策）
                        // 一時ファイルの場合は登録しない
                        if ((Path.GetExtension(strFilePath) == exp) &&
                            (Path.GetFileName(strFilePath).Substring(0, 2) != "~$"))
                        {
                            listFiles.Add(strFilePath);
                        }
                    }

                    // 画面に進捗表示
                    labelStatus.Text = string.Format(Resources.msgStatusSearch, listFiles.Count);   // step2    iwasa

                }

                // 検索したファイルリストをstring配列に変換
                stFileList = listFiles.ToArray();
            }
            catch
            {
                ProcControlEnabled(true);

                MessageBox.Show(Resources.msgFailedSearchFile, // step2    iwasa
                         Resources.msgError,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Hand);

                return;
            }

            #endregion

            int count = 0;
            #region<内部保持用ファイルリスト登録処理>
            try
            {
                // 内部保持用ファイルリストクリア
                ArrayFileList.Clear();

                // ハッシュリストクリア
                htList.Clear();

                // ファイル詳細情報取得
                for (int i = 0; i < stFileList.Count(); i++)
                {
                    // 待機
                    Delay(0.01);

                    // 中止ボタン押下した場合は処理終了
                    if (bStopFlg == true) break;

                    // 進捗表示
                    labelStatus.Text = string.Format(Resources.msgStatusInfo, i + 1, stFileList.Count());   // step2    iwasa

                    // 内部保持用ファイルリストにデータ追加
                    SetFileListAdd(stFileList[i], ref ArrayFileList, ref fileReadErrFlg);  // 20171012 修正（ファイル読込時にエラーが発生した場合にフラグを付ける）

                    // ハッシュテーブルに追加
                    htList.Add(stFileList[i], stFileList[i]);
                    count++;
                }

                // 絞込みおよびリスト表示処理(絞込みボタン押下処理)
                buttonRefine_Click(sender, e);
            }
            catch
            {
                ProcControlEnabled(true);

                MessageBox.Show(Resources.msgFailedFileInfo, // step2    iwasa
                         Resources.msgError,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Hand);
                return;
            }
            #endregion

            // 20170731 追加（システムファイルにアクセスした場合に表示）
            if (accessSystemFileFlg == true)
            {
                string systemfilepath = Path.Combine(EXCLUSION_LOG_PATH, "Log");
                MessageBox.Show(Resources.msgAccessSystemFile + systemfilepath, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }

            // 画面入力可
            ProcControlEnabled(true);
        }

        /// <summary>
        ///  20170731 追加（システムファイルの読込時にシステムファイル以外の検索結果は表示するよう修正）
        /// </summary>
        /// <param name="path"></param>
        /// <param name="error_dt"></param>
        /// <param name="error_list"></param>
        /// <param name="isSubDirectory">サブフォルダを含めるフラグ</param>
        /// <returns></returns>
        List<string> GetNotSystemFileList(string path, string error_dt, ref List<string> error_list , bool isSubDirectory)
        {
            string[] subFolders = { };
            string[] subFiles = { };

            // システムフォルダ/ファイル読込時は「catch」に入る
            try
            {
                // 指定フォルダ配下のフォルダを全て取得（サブフォルダは含まない）
                subFolders = System.IO.Directory.GetDirectories(path, "*", System.IO.SearchOption.TopDirectoryOnly);
                // 指定フォルダ配下のファイルを全て取得（サブフォルダは含まない）
                subFiles = System.IO.Directory.GetFiles(path, "*", System.IO.SearchOption.TopDirectoryOnly);
            }
            catch(Exception e)
            {
                accessSystemFileFlg = true;

                // ログ出力リストに追加
                error_list.Add(path);
            }

            // 返却値（最終的な返り値）
            List<string> ret = new List<string>();
            ret.AddRange(subFiles);

            if(isSubDirectory == true)
            {
                // 指定フォルダ配下のフォルダを検索（再帰）
                for (int i = 0; i < subFolders.Length; i++)
                {
                    // 返却値2（再起で使用）
                    List<string> ret_2 = new List<string>();
                    ret_2 = GetNotSystemFileList(subFolders[i], error_dt, ref error_list, isSubDirectory);
                    ret.AddRange(ret_2);
                }
            } 
            
            return ret;
        }

        /// <summary>
        /// 絞込ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRefine_Click(object sender, EventArgs e)
        {
            CurrentArrayGridOutput.Clear();
            ArrayList arrayGridOutput = CurrentArrayGridOutput;

            #region<検索条件取得>
            // 検索ファイル名取得
            string strSearchName = textBoxFileName.Text;

            // 作成日範囲取得
            DateTime? dtCreateFrom = null, dtCreateTo = null;
            if (checkBoxCreate.Checked)
            {
                dtCreateFrom = new DateTime(dateTimePickerCreateFrom.Value.Year, dateTimePickerCreateFrom.Value.Month, dateTimePickerCreateFrom.Value.Day, 0, 0, 0);
                dtCreateTo = new DateTime(dateTimePickerCreateTo.Value.Year, dateTimePickerCreateTo.Value.Month, dateTimePickerCreateTo.Value.Day, 0, 0, 0).AddDays(1);
            }

            // 更新日範囲取得
            DateTime? dtUpdateFrom = null, dtUpdateTo = null;
            if (checkBoxUpdate.Checked)
            {
                dtUpdateFrom = new DateTime(dateTimePickerUpdateFrom.Value.Year, dateTimePickerUpdateFrom.Value.Month, dateTimePickerUpdateFrom.Value.Day, 0, 0, 0);
                dtUpdateTo = new DateTime(dateTimePickerUpdateTo.Value.Year, dateTimePickerUpdateTo.Value.Month, dateTimePickerUpdateTo.Value.Day, 0, 0, 0).AddDays(1);
            }

            // ファイルタイプ
            if (radioButtonTypeAll.Checked == true)
            {
                IListFileType = EXTENSION_TYPE_ALL; // step2 iwasa
            }
            else if (radioButtonTypeOfficePdf.Checked == true)
            {
                IListFileType = EXTENSION_TYPE_OFFICEPDF;
            }
            else if (radioButtonTypeOffice.Checked == true)
            {
                IListFileType = EXTENSION_TYPE_OFFICE;
            }
            else
            {
                IListFileType = EXTENSION_TYPE_PDF;
            }
            //IListFileType = (radioButtonTypeOfficePdf.Checked == true) ? EXTENSION_TYPE_OFFICEPDF : (radioButtonTypeOffice.Checked == true) ? EXTENSION_TYPE_OFFICE : EXTENSION_TYPE_PDF;
            #endregion

            #region<リスト絞込処理>
            foreach (GridListData clsListData in ArrayFileList)
            {
                // 絞込条件に一致しない場合は登録しない
                if (SetFileListRefine(clsListData, IListFileType, strSearchName, dtCreateFrom, dtCreateTo, dtUpdateFrom, dtUpdateTo, true, true))
                {
                    // 表示用リストに追加
                    arrayGridOutput.Add(clsListData);
                }
            }
            #endregion

            // リスト画面表示
            dataGridView_redraw(arrayGridOutput);

            // ファイルパスの部分を自動で幅を調整する
            dataGridViewList.AutoResizeColumn((int)DataGridViewColumnIndex.FilePath, DataGridViewAutoSizeColumnMode.AllCells);
        }

        /// <summary>
        /// リスト ドラッグアンドドロップ処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewList_DragDrop(object sender, DragEventArgs e)
        {
            #region<データ追加確認ダイアログ表示処理>

            DialogResult dr = MessageBox.Show(
                        Resources.msgDragAndDrop,
                        Resources.msgConfirmation,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Exclamation);

            if (dr == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            #endregion

            ArrayList arrayFileList = new ArrayList();          // 表示用ファイルリスト

            bool fileReadErrFlg = false;                          // 20171012 追加 ファイル読込時にエラーが発生した場合:true

            // 画面入力不可
            ProcControlEnabled(false);

            // 進捗表示
            //labelStatus.Text = "ファイル検索中 対象ファイル 0件";
            labelStatus.Text = string.Format(Resources.msgStatusSearch, 0);   // step2    iwasa

            #region<ファイル検索処理>
            List<string> listFiles = new List<string>();
            Boolean bConfirm = false;
            System.IO.SearchOption searchOption = (radioButtonSubFolderInc.Checked) ? System.IO.SearchOption.AllDirectories : System.IO.SearchOption.TopDirectoryOnly;

            // 20200804 zipファイル対応
            string error_dt = DateTime.Now.ToString("yyyyMMddHHmmss");  // エラーファイルに使用する日付
            List<string> error_log = new List<string>();
            List<string> ZipBufFiles = new List<string>();

            // ファイル詳細情報絞込み
            foreach (string strDDFilePath in (string[])e.Data.GetData(DataFormats.FileDrop))
            {
                if (System.IO.File.Exists(strDDFilePath) == true)
                {
                    // ファイルがドロップされた場合
                    if (Path.GetExtension(strDDFilePath).Contains("zip") != false)
                    {
                        ZipBufFiles.Add(strDDFilePath);
                    }

                    // リストに同じパスのファイルが登録されていない場合のみ登録する
                    if (htList.ContainsValue(strDDFilePath) == false)
                    {
                        listFiles.Add(strDDFilePath);
                    }

                    // 画面に進捗表示
                    labelStatus.Text = string.Format(Resources.msgStatusSearch, listFiles.Count);   // step2    iwasa
                }
                else if (System.IO.Directory.Exists(strDDFilePath) == true)
                {
                    // フォルダがドロップされた場合
                    ZipBufFiles = GetNotSystemFileList(strDDFilePath, error_dt, ref error_log, true);
                    ZipBufFiles.RemoveAll(CompressedFileChecker.judge);

                    foreach (string exp in StrExtension_Narrow)
                    {
                        // 待機
                        Delay(0.01);

                        // 中止ボタン押下した場合は処理終了
                        if (bStopFlg == true) break;

                        // 拡張毎にファイルリストを取得
                        string[] listBufFiles = System.IO.Directory.GetFiles(strDDFilePath, "*" + exp, searchOption);


                        foreach (string strFilePath in listBufFiles)
                        {
                            // 指定の拡張子以外の場合は登録しない（*.xlsで検索すると.xlsx、.xlsmが検索されてしまう対策）
                            if (Path.GetExtension(strFilePath) == exp)
                            {
                                // リストに同じパスのファイルが登録されていない場合のみ登録する
                                if (htList.ContainsValue(strFilePath) == false)
                                {
                                    listFiles.Add(strFilePath);
                                }
                            }
                        }

                        // 画面に進捗表示
                        labelStatus.Text = string.Format(Resources.msgStatusSearch, listFiles.Count);   // step2    iwasa

                        // 指定件数を超えた場合は続行の確認ダイアログを表示する
                        if ((!bConfirm) && (listFiles.Count > SEARCH_ALERT_COUNT))
                        {
                            DialogResult drWarning = MessageBox.Show(
                                Resources.msgFileTarget + SEARCH_ALERT_COUNT + Resources.msgFileContinueProcess,
                                Resources.msgConfirmation,
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Exclamation);

                            if (drWarning == System.Windows.Forms.DialogResult.No)
                            {
                                // 画面入力可
                                ProcControlEnabled(true);
                                return;
                            }
                            bConfirm = true;
                        }
                    }
                }
            }

            bool IsZipTarget = checkBoxZipTarget.Checked;

            // Zipファイルの登録
            CompressedFileChecker zipTest = new CompressedFileChecker();

            if (IsZipTarget != false)
            {
                zipTest.settingForm = settingForm;
                zipTest.GetZipAllList(ZipBufFiles, ref dicZipResult, ref dicCompressedItem, ref dicPasswordZip, ref dicErrorZip);

                foreach (var Files in ZipBufFiles)
                {
                    foreach (string path in dicCompressedItem[Files])
                    {
                        // 拡張子の数だけループ
                        foreach (string exp in StrExtension_Narrow)
                        {
                            // 指定の拡張子以外の場合は登録しない（*.xlsで検索すると.xlsx、.xlsmが検索されてしまう対策）
                            // 一時ファイルの場合は登録しない
                            if ((Path.GetExtension(path) == exp) &&
                                (Path.GetFileName(path).Substring(0, 2) != "~$"))
                            {
                                // リストに同じパスのファイルが登録されていない場合のみ登録する
                                if (htList.ContainsValue(path) == false)
                                {
                                    listFiles.Add(path);
                                }
                            }
                        }
                    }

                }
            }

            // 検索したファイルリストをstring配列に変換
            string[] stFileList = listFiles.ToArray();
            #endregion

            #region<内部保持用ファイルリスト登録処理>
            // ファイル詳細情報取得
            for (int i = 0; i < stFileList.Count(); i++)
            {
                // 待機
                Delay(0.01);

                // 中止ボタン押下した場合は処理終了
                if (bStopFlg == true) break;

                // 進捗表示
                labelStatus.Text = string.Format(Resources.msgStatusInfo, i + 1, stFileList.Count());

                // 内部保持用ファイルリストにデータ追加
                SetFileListAdd(stFileList[i], ref ArrayFileList, ref fileReadErrFlg);

                // ハッシュテーブルに追加
                htList.Add(stFileList[i], stFileList[i]);
            }
            #endregion

            // 絞込みおよびリスト表示処理(絞込みボタン押下処理)
            buttonRefine_Click(sender, e);

            // 20171012 修正（ファイル読込時にエラーが発生した場合にダイアログを表示）
            if (fileReadErrFlg)
            {
                MessageBox.Show(Resources.msgFailedReadPropertyFile, // step2    iwasa
                         Resources.msgError,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Hand);
            }

            // 画面入力可
            ProcControlEnabled(true);
        }

        /// <summary>
        /// リスト ドラッグ処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewList_DragEnter(object sender, DragEventArgs e)
        {
            // ファイルのドラッグアンドドロップのみを受け付ける
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // ドロップファイルをアプリケーション側へ内容をコピー
                e.Effect = DragDropEffects.Copy;
            }
        }

        /// <summary>
        /// リストマウスダウン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewList_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            // 右クリック時の場合
            if (e.Button == MouseButtons.Right)
            {
                if (dataGridViewList.Rows.Count > 0)
                {
                    if ((e.RowIndex >= 0) && (e.ColumnIndex >= 0))
                    {
                        // 選択状態をクリア
                        dataGridViewList.ClearSelection();

                        // クリック位置を選択状態にする
                        dataGridViewList.Rows[e.RowIndex].Selected = true;

                        // 右クリックメニュー表示
                        System.Drawing.Point p = System.Windows.Forms.Cursor.Position;
                        this.contextMenuStrip.Show(p);
                    }
                }
            }
        }

        /// <summary>
        /// リスト描画処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewList_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // 列ヘッダーのみ処理を行う。(CheckBox配置列が先頭列の場合)
            if (e.ColumnIndex == (int)DataGridViewColumnIndex.Select && e.RowIndex == -1)
            {
                using (Bitmap bmp = new Bitmap(100, 100))
                {
                    // チェックボックスの描画領域を確保
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.Clear(Color.Transparent);
                    }

                    // 描画領域の中央に配置
                    Point pt1 = new Point(5, (bmp.Height - checkBoxAll.Height) / 2);
                    if (pt1.Y < 0) pt1.Y = 0;

                    // Bitmapに描画
                    checkBoxAll.DrawToBitmap(bmp, new System.Drawing.Rectangle(pt1.X, pt1.Y, bmp.Width, bmp.Height));

                    // DataGridViewの現在描画中のセルの中央に描画
                    int x = 0;
                    int y = (e.CellBounds.Height - bmp.Height) / 2;
                    Point pt2 = new Point(e.CellBounds.Left + x, e.CellBounds.Top + y);

                    e.Paint(e.ClipBounds, e.PaintParts);
                    e.Graphics.DrawImage(bmp, pt2);
                    e.Handled = true;
                }
            }

            // zipファイルと解凍不可はチェックボックスを灰色にする step2 iwasa
            if (e.ColumnIndex == (int)DataGridViewColumnIndex.Select && e.RowIndex != -1)
            {
                DataGridViewCell cell = dataGridViewList[e.ColumnIndex, e.RowIndex];
                DataGridViewCheckBoxCell checkCell = cell as DataGridViewCheckBoxCell;

                // zip or 解凍不可のとき
                if (checkCell.ReadOnly == true)
                {
                    bool _selected = (e.State == DataGridViewElementStates.Selected);

                    e.PaintBackground(e.CellBounds, _selected);

                    int _size = 14;
                    Rectangle rect = e.CellBounds;

                    rect.Width = _size;
                    rect.Height = _size;
                    rect.Offset((e.CellBounds.Width - _size) / 2, rect.Height / 2 - 4);

                    ControlPaint.DrawMixedCheckBox(e.Graphics, rect, ButtonState.Inactive);
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// リストセルクリック時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // ヘッダーの選択チェックボックス選択時
            if (e.ColumnIndex == (int)DataGridViewColumnIndex.Select && e.RowIndex == -1)
            {
                checkBoxAll.Checked = !checkBoxAll.Checked;
            }

            // 選択チェックボックス選択時
            if (e.ColumnIndex == (int)DataGridViewColumnIndex.Select && e.RowIndex >= 0)
            {
                // 選択列のセルクリックでチェックボックス選択
                DataGridViewCell cell = dataGridViewList[e.ColumnIndex, e.RowIndex];
                DataGridViewCheckBoxCell checkCell = cell as DataGridViewCheckBoxCell;
                // ReadOnlyがtrueならチェックを入れない
                if(checkCell.ReadOnly == true)
                {
                    checkCell.Value = checkCell.FalseValue;
                }
                else
                {
                    checkCell.Value = (checkCell.Value == checkCell.TrueValue) ? checkCell.FalseValue : checkCell.TrueValue;
                }

                // カーソルを移動
                dataGridViewList.CurrentCell = dataGridViewList[(int)DataGridViewColumnIndex.UpdatedDate, e.RowIndex];

                // 選択されている項目チェック
                int iSelectedCount = 0;
                foreach (DataGridViewRow row in dataGridViewList.Rows)
                {
                    // チェックしている項目の場合
                    if (row.Cells[(int)DataGridViewColumnIndex.Select].Value.ToString() == "1") iSelectedCount++;
                }
                // 全ての行にチェックがついている場合は全選択チェックボックスをチェックする。そうでない場合はチェックを外す。(割り込み禁止使用)
                bInterruptDisabled = true;
                checkBoxAll.Checked = (iSelectedCount == dataGridViewList.Rows.Count) ? true : false;
                bInterruptDisabled = false;
                dataGridViewList.Refresh();
            }
        }

        /// <summary>
        /// リストセルダブルクリック時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewList_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // 選択チェックボックス選択時
            if (e.ColumnIndex == (int)DataGridViewColumnIndex.Select && e.RowIndex >= 0)
            {
                // カーソルを移動
                dataGridViewList.CurrentCell = dataGridViewList[(int)DataGridViewColumnIndex.UpdatedDate, e.RowIndex];
            }
        }

        /// <summary>
        /// ヘッダーの選択チェックボックス選択時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxAll_CheckedChanged(object sender, EventArgs e)
        {
            // 割り込み禁止でない場合
            if (!bInterruptDisabled)
            {
                // 全チェックボックス更新
                foreach (DataGridViewRow row in this.dataGridViewList.Rows)
                {
                    // ReadOnlyがfalseのチェックボックスだけチェックを付ける
                    if (row.Cells[(int)DataGridViewColumnIndex.Select].ReadOnly == false)
                    {
                        row.Cells[(int)DataGridViewColumnIndex.Select].Value = (checkBoxAll.Checked) ? "1" : "0";
                    }
                }
            }
        }

        /// <summary>
        /// キー入力
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewList_KeyDown(object sender, KeyEventArgs e)
        {
            // Deleteキーが押された場合
            if (e.KeyData == Keys.Delete)
            {
                //リストで選択されている行を削除する
                foreach (DataGridViewRow r in dataGridViewList.SelectedRows)
                {
                    if (!r.IsNewRow)
                    {
                        // リストで選択されている行のファイルパスを取得
                        string IsZipFormat = r.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString();
                        if (string.IsNullOrEmpty(IsZipFormat) == false)
                        {
                            // ZIP本体またはZIPファイルの中身の場合

                            DialogResult dr = MessageBox.Show(
                                Resources.msgZipListDelete,
                                Resources.msgConfirmation,
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Exclamation);

                            if (dr == System.Windows.Forms.DialogResult.OK)
                            {
                                // はい

                                // zipとそれに入っているファイルをリストからすべて削除

                                string CommonZipPath = r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString();

                                List<DataGridViewRow> listRow = new List<DataGridViewRow>();
                                foreach (DataGridViewRow targetRow in dataGridViewList.Rows)
                                {
                                    if (targetRow.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value != null)
                                    {
                                        string targetPath = targetRow.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString();
                                        if (targetPath == CommonZipPath)
                                        {
                                            // 同一ZIPデータ
                                            listRow.Add(targetRow);
                                        }
                                    }
                                }

                                foreach (DataGridViewRow delRow in listRow)
                                {
                                    // ハッシュテーブルから削除
                                    htList.Remove(delRow.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString());

                                    // 内部保持用ファイルリストから削除
                                    for (int i = 0; i < ArrayFileList.Count; i++)
                                    {
                                        GridListData clsListData = (GridListData)ArrayFileList[i];
                                        if (clsListData.strTmpZipFilePath == delRow.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString())
                                        {
                                            ArrayFileList.RemoveAt(i);
                                            break;
                                        }
                                    }

                                    // リストから削除
                                    dataGridViewList.Rows.Remove(delRow);
                                }
                            }
                            // キャンセル
                            return;
                        }


                        // ハッシュテーブルから削除
                        htList.Remove(r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString());

                        // 内部保持用ファイルリストから削除
                        for (int i = 0; i < ArrayFileList.Count; i++)
                        {
                            GridListData clsListData = (GridListData)ArrayFileList[i];
                            if (clsListData.strFilePath == r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString())
                            {
                                ArrayFileList.RemoveAt(i);
                                break;
                            }
                        }

                        // リストから削除
                        dataGridViewList.Rows.Remove(r);
                    }
                }
            }
        }

        /// <summary>
        /// 設定開始ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonStart_Click(object sender, EventArgs e)
        {
            int iSelectedCount = 0;                                 // 選択されている数
            int iErrorCount = 0;                                    // エラー件数

            bool confirm_skip = false;
            #region<件数チェック>
            // データがない場合
            if (dataGridViewList.Rows.Count == 0)
            {
                MessageBox.Show(Resources.msgNotFileExist, Resources.msgConfirmation, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 選択されている項目チェック、設定済みチェック
            Boolean bSettingFlg = false;                            // 属性情報設定ありフラグ
            foreach (DataGridViewRow row in dataGridViewList.Rows)
            {
                // チェックしている項目の場合
                if (row.Cells[(int)DataGridViewColumnIndex.Select].Value.ToString() == "1")
                {
                    // 選択されている数加算
                    iSelectedCount++;

                    // 属性情報設定があるかチェック
                    if (((row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value == null) || (row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value.ToString() != "")) &&
                         ((row.Cells[(int)DataGridViewColumnIndex.SecretType].Value == null) || (row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString() != "")))
                    {
                        // 属性情報設定ありフラグON
                        bSettingFlg = true;
                    }
                }
            }

            // 選択項目がない場合
            if (iSelectedCount == 0)
            {
                MessageBox.Show(Resources.msgSelectData, Resources.msgConfirmation, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 属性情報設定ありフラグONの場合は確認ダイアログを表示する設定できなかった
            if (bSettingFlg)
            {
                DialogResult dr = MessageBox.Show(
                        Resources.msgChangeSetting,
                        Resources.msgConfirmation,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Exclamation);

                if (dr == System.Windows.Forms.DialogResult.No)    // step2    iwasa
                {
                    return;
                }
            }

            #endregion

            // 画面入力不可
            ProcControlEnabled(false);

            // 進捗表示
            labelStatus.Text = string.Format(Resources.msgStatusSetting);   // step2    iwasa

            #region<属性情報変更>

            // 各フォーム呼び出し
            SettingForm clsHliSetting = new SettingForm();
            clsHliSetting.clsCommonSettting = settingForm.clsCommonSettting;

            // リスト表示分繰り返す
            int iProcessCount = 1;                                  // 処理件数
            string strClassNo = "";                                 // 文書分類番号
            string strSecrecyLevel = "";                            // 機密区分

            DialogResult RankDownResult = new DialogResult();

            // 選択したファイルがzip内のファイルの場合に管理するリスト
            HashSet<string> listZipTarget = new HashSet<string>();

            foreach (DataGridViewRow row in dataGridViewList.Rows)
            {
                // 中止ボタン押下した場合は処理終了
                if (bStopFlg == true) break;

                // 待機
                Delay(0.01);

                // チェックしていない場合はスキップ
                if (row.Cells[(int)DataGridViewColumnIndex.Select].Value.ToString() != "1") continue;

                // 拡張子取得
                // 20170905 修正
                int iExtension = ExtensionCheck(row.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString() + @"\" + row.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString());

                // 指定拡張子以外の場合はスキップ
                if (iExtension == EXTENSION_NONE) continue;

                // 行選択
                row.Selected = true;
                dataGridViewList.FirstDisplayedScrollingRowIndex = row.Index;

                // ファイルタイプ取得
                int iListFileType = (iExtension == EXTENSION_TYPE_OFFICEPDF) ? EXTENSION_TYPE_OFFICEPDF : (iExtension == EXTENSION_PDF) ? EXTENSION_TYPE_PDF : EXTENSION_TYPE_OFFICE;

                // 「一件毎に機密区分・文書分類指定」にチェックが入っている場合、または属性設定が設定されていない場合はダイアログを表示する
                if ((radioButtonDesignation.Checked == true) || ((strSecrecyLevel == "")))
                {

                    // 一括設定時のデフォルト値を設定値に変更
                    clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel = (radioButtonBatch.Checked) ? settingForm.clsCommonSettting.strDefaultSecrecyLevel : row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString();

                    System.Windows.Forms.DialogResult dlgResult = clsHliSetting.ShowDialog();
                    if (dlgResult == System.Windows.Forms.DialogResult.No)
                    {
                        // 閉じるボタン押下の場合
                        break;
                    }
                    else if (dlgResult == System.Windows.Forms.DialogResult.OK)
                    {
                        if (clsHliSetting.IsStampProc == false)
                        {
                            if (radioButtonBatch.Checked != false)
                            {
                                // 一括
                                foreach (DataGridViewRow CheckRow in dataGridViewList.Rows)
                                {
                                    string SettingSecrecy = clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel;
                                    string CurrentSecrecy = CheckRow.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString();
                                    if (CurrentSecrecy != SettingSecrecy)
                                    {
                                        MessageBox.Show(Resources.msgStampOfficeFile, Resources.msgConfirmation, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        clsHliSetting.IsStampProc = true;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                // 個別
                                string SettingSecrecy = clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel;
                                string CurrentSecrecy = row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString();
                                if (CurrentSecrecy != SettingSecrecy)
                                {
                                    MessageBox.Show(Resources.msgStampOfficeFile, Resources.msgConfirmation, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    clsHliSetting.IsStampProc = true;
                                }
                            }
                        }

                        // 20200807
                        strSecrecyLevel = clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel;
                    }
                    else
                    {
                        // 後で登録ボタン押下の場合

                        // 一括変更の場合は処理終了
                        if (radioButtonBatch.Checked) break;

                        // 登録設定なし
                        strClassNo = "";
                        strSecrecyLevel = "";
                    }
                }

                // 20200804 属性を見て処理中断
                string SecureTarget = row.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString();
                if (string.IsNullOrEmpty(row.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString()) != true)
                {
                    // zipファイルの中身の場合
                    SecureTarget = Path.GetDirectoryName(row.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString());
                }

                List<string> lstTarGetSecureFolder = settingForm.clsCommonSettting.lstSecureFolder;
                if ((clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel == SECRECY_PROPERTY_S) ||
                    (clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel == SECRECY_PROPERTY_A)
                    )
                {
                    string result = lstTarGetSecureFolder.FirstOrDefault(x => SecureTarget.Contains(x));
                    if (result == null)
                    {
                        // セキュアフォルダではない
                        string NotSecureFolderMsg = row.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString() + Resources.msgNotSecureFolder;
                        MessageBox.Show(NotSecureFolderMsg, Resources.msgConfirmation, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);   
                        continue;
                    }
                }

                // 属性設定処理
                if (strSecrecyLevel != "")
                {
                    // 現在のランクがSで、変更後がAかBの場合 または、現在のランクがAで、変更後がBの場合
                    if (row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString() == SECRECY_PROPERTY_S && (clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel == SECRECY_PROPERTY_A || clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel == SECRECY_PROPERTY_B)
                          || (row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString() == SECRECY_PROPERTY_A && clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel == SECRECY_PROPERTY_B))

                    {
                        // 今後このメッセージを表示しないがチェックされていない場合
                        if (!confirm_skip)
                        {
                            // ランク降格の確認ダイアログを表示
                            RankDownConfirm rankdownConfirm = new RankDownConfirm(row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString(), clsHliSetting.clsCommonSettting.strDefaultSecrecyLevel, radioButtonBatch.Checked);

                            rankdownConfirm.StartPosition = FormStartPosition.CenterParent;
                            rankdownConfirm.ShowDialog();
                            // 今後このメッセージを表示しないフラグを取得
                            confirm_skip = rankdownConfirm.bMsgSkip;
                            RankDownResult = rankdownConfirm.DialogResult;
                        }

                        if (RankDownResult == DialogResult.OK)
                        {
                            // 書き込む(後続処理を継続)
                        }
                        else if(RankDownResult == DialogResult.Cancel)
                        {
                            // 中止
                            break;
                        }
                        else if (RankDownResult == DialogResult.Ignore) // DialogResultにスキップ変数が無いのでignoreで代用
                        {
                            // スキップ
                            continue;
                        }
                    }

                    string file_name = row.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString() + @"\" + row.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString();

                    if (string.IsNullOrEmpty(row.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString()) == false)
                    {
                        // zipファイルの中身の場合
                        file_name = row.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString();
                    }

                    string keyword;
                    DateTime updateTime = new DateTime();


                    // 20171009 修正（第3パラメータが「無」or「no_stamp」の場合は「HM」に付替）
                    // 空欄の場合自身の事業者コードにする
                    keyword = strSecrecyLevel + SEPARATE + " " + SEPARATE + " " + clsHliSetting.clsCommonSettting.strOfficeCode + SEPARATE;

                    Boolean bPropertyResult = false;                        // プロパティ設定結果
                    string error_reason = "";                               // 20171011 追加（ファイル書込み時にエラーが発生した場合にエラー理由を保存）
                    if (iListFileType == EXTENSION_TYPE_PDF)
                    {
                        // PDFファイルの場合
                        // 20171011 追加（ファイル書込み時にエラーが発生した場合にエラー理由を保存）
                        // 20171213 TSE demachi CHG 引数にupdateTimeを追加 更新日が入り、グリッド更新に使う
                        bPropertyResult = WriteByPDF(file_name, keyword, false, ref updateTime, ref error_reason);
                    }
                    else
                    {
                        bool bOpenXML = ExtensionOpenXMLCheck(file_name);

                        string beforeSecrecyLevel = (string)row.Cells[(int)DataGridViewColumnIndex.SecretType].Value;
                        bool IsSuccess = clsHliSetting.SetStamp(iExtension, file_name, beforeSecrecyLevel);

                        if (IsSuccess != false)
                        {
                            if (bOpenXML)
                            {
                                // Officeファイルの場合
                                // 20170905 TSE kitada OFFICE2016対応
                                AccessToProperties atp = new AccessToProperties();
                                bPropertyResult = atp.WriteProperties(file_name, keyword, iExtension, false, clsHliSetting.clsCommonSettting.lstFinal,   // step2 iwasa
                                                                      ref updateTime, ref error_reason);  // 20171011 追加（ファイル書込み時にエラーが発生した場合にエラー理由を保存）
                            }
                            else
                            {
                                bPropertyResult = WriteByDSO(file_name, keyword, false, ref updateTime, ref error_reason);  // 20171011 追加（ファイル書込み時にエラーが発生した場合にエラー理由を保存）
                            }
                        }
                    }

                    // 更新に成功した場合は画面更新
                    if (bPropertyResult)
                    {
                        // リスト内更新処理 チェック、文書分類、機密区分
                        row.Cells[(int)DataGridViewColumnIndex.UpdatedDate].Value = updateTime.ToString("yyyy/MM/dd HH:mm:ss");    // 20171027 更新日を更新
                        row.Cells[(int)DataGridViewColumnIndex.Select].Value = "0";
                        row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value = strClassNo;  // 20171009 修正（他事業所ファイルの場合は事業所名を表示）
                        row.Cells[(int)DataGridViewColumnIndex.SecretType].Value = strSecrecyLevel;
                        row.Cells[(int)DataGridViewColumnIndex.SecretTypeHidden].Value = strClassNo;       // 20170928 TSE kitada 追加　機密分類_hidden

                        // 内部リストの内容を更新する
                        for (int i = 0; i < ArrayFileList.Count; i++)
                        {
                            GridListData clsListData = (GridListData)ArrayFileList[i];

                            string GridListFileName = row.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString() + @"\" + row.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString();
                            string strFilePath = clsListData.strFilePath;

                            if (row.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value != null)
                            {
                                GridListFileName = row.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString();
                                strFilePath = clsListData.strTmpZipFilePath;

                            }

                            // 20200805 修正
                            if (strFilePath == GridListFileName)
                            {
                                if (Path.GetExtension(clsListData.strFilePath).Contains("zip") != false)
                                {
                                    listZipTarget.Add(clsListData.strFilePath);
                                }

                                clsListData.strClassNo = strClassNo;
                                clsListData.strSecrecyLevel = strSecrecyLevel;
                                ArrayFileList[i] = clsListData;
                                break;
                            }
                        }

                        // 全選択チェックがついている場合はチェックをはずす
                        if (checkBoxAll.Checked)
                        {
                            bInterruptDisabled = true;
                            checkBoxAll.Checked = false;
                            bInterruptDisabled = false;
                            Delay(0.01);
                        }
                    }
                    else
                    {
                        // エラー件数加算
                        iErrorCount++;

                        // リストの分類表示をエラー表示にする
                        row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value = error_reason;    // 20171011 修正（エラー理由によって表示文言を変更）
                    }
                }

                // 進捗表示
                labelStatus.Text = string.Format(Resources.msgStatusInfo, iProcessCount, iSelectedCount);

                // 処理件数加算
                iProcessCount++;
            }
            #endregion

            // 再zip
            CompressedFileChecker zipUtil = new CompressedFileChecker();
            zipUtil.settingForm = settingForm;
            zipUtil.SelectZipProc(ref listZipTarget, ref dicZipResult);

            // 画面入力可
            ProcControlEnabled(true);

            // エラー時の場合はメッセージを表示する
            if (iErrorCount > 0)
            {
                // 20171011 修正（エラー理由によって表示文言を変更）    // step2    iwasa
                string errorcountMsg = string.Format("{0}" + Resources.msgFailedSettingFile, iErrorCount);
                MessageBox.Show(errorcountMsg, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand); 
            }

            // 先頭行までスクロール
            dataGridViewList.FirstDisplayedScrollingRowIndex = 0;
        }

        /// <summary>
        /// 中止ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonStop_Click(object sender, EventArgs e)
        {
            // 中止フラグON
            bStopFlg = true;
        }

        /// <summary>
        /// Excel出力ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonExcelOutput_Click(object sender, EventArgs e)
        {
            try
            {
                #region<保存ダイアログ表示>
                SaveFileDialog sfd = new SaveFileDialog();              // 保存ダイアログ

                // 初期フォルダを指定する
                sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // タイトルを設定する
                sfd.Title = Resources.msgSelectFile;    // step2    iwasa

                // [ファイルの種類]に表示される選択肢を指定する
                sfd.Filter = Resources.msgExcelfile;    // step2    iwasa

                // 20170913追加 （デフォルトのファイル名）
                sfd.FileName = Resources.msgExportExcelDefault + DateTime.Now.ToString("yyyyMMddHHmmss");

                // 保存ダイアログを表示する
                if (sfd.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                #endregion

                #region<Excelファイルに書き込み>
                // 描画停止
                dataGridViewList.SuspendLayout();

                dataGridViewList.Visible = false;

                // 現在のソート条件を取得
                DataGridViewColumn sortColumn = (dataGridViewList.SortedColumn == null) ? dataGridViewList.Columns[(int)DataGridViewColumnIndex.FilePath] : dataGridViewList.SortedColumn;

                ListSortDirection sortDirection = (dataGridViewList.SortOrder == SortOrder.Descending) ? ListSortDirection.Descending : ListSortDirection.Ascending;

                // ファイルパス(昇順)でソートを行う
                dataGridViewList.Sort(dataGridViewList.Columns[(int)DataGridViewColumnIndex.FilePath], ListSortDirection.Ascending);

                // ファイルを作成
                Excel.Application ExcelApp = new Excel.Application();

                // Excelウィンドウを非表示する
                ExcelApp.Visible = false;

                // アクティブのワークブックを取得
                Excel.Workbook wb = ExcelApp.Workbooks.Add();

                // 先頭のシートを取得
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];

                #region<リストからオブジェクトリストに登録>
                // stringリストにデータを格納
                string[,] strs = new string[dataGridViewList.RowCount + 1, EXCEL_OUTPUT_COLUMN_COUNT];
                string[] headerTitle = new string[EXCEL_OUTPUT_COLUMN_COUNT]    // step2 iwasa
                                                 { Resources.msgZip, Resources.msgFileName, Resources.msgCreateDate, Resources.msgUpdateDate,
                                                   Resources.msgDocumentClass, Resources.msgClassified, Resources.msgFilePath, Resources.msgAuthor, Resources.msgLastUpdate};

                for(int i=0; i< EXCEL_OUTPUT_COLUMN_COUNT; i++)
                {
                    strs[0, i] = headerTitle[i];
                }

                for (int i = 0; i < dataGridViewList.RowCount; i++)
                {
                    int[] column = new int[EXCEL_OUTPUT_COLUMN_COUNT]   // step2 iwasa
                                          {(int)DataGridViewColumnIndex.ZipFormat,   (int)DataGridViewColumnIndex.FileName,     (int)DataGridViewColumnIndex.CreatedDate,
                                           (int)DataGridViewColumnIndex.UpdatedDate, (int)DataGridViewColumnIndex.DocumentType, (int)DataGridViewColumnIndex.SecretType,
                                           (int)DataGridViewColumnIndex.FilePath,    (int)DataGridViewColumnIndex.CreatedBy,    (int)DataGridViewColumnIndex.UpdatedBy};

                    for(int j=0; j< EXCEL_OUTPUT_COLUMN_COUNT; j++)
                    {
                        strs[i + 1, j] = dataGridViewList.Rows[i].Cells[column[j]].Value.ToString();
                    }
                }

                // 元の順にソートし直す
                dataGridViewList.Sort(sortColumn, sortDirection);

                // 描画開始
                dataGridViewList.Visible = true;
                
                dataGridViewList.ResumeLayout();

                // オブジェクトリストにデータを関連付け
                object[,] datas = new object[dataGridViewList.RowCount + 1, EXCEL_OUTPUT_COLUMN_COUNT];

                for (int i = 0; i < dataGridViewList.RowCount + 1; i++)
                {
                    for (int j = 0; j < EXCEL_OUTPUT_COLUMN_COUNT; j++)
                    {
                        datas[i, j] = strs[i, j];
                    }
                }
                #endregion

                // 書式設定を文字列にする  step2 iwasa
                int[] strFormat = new int[] { EXCEL_SHEET_COLUMN1,    // zip形式
                                              EXCEL_SHEET_COLUMN2,    // ファイル名
                                              EXCEL_SHEET_COLUMN5,    // 文書分類
                                              EXCEL_SHEET_COLUMN6,    // 機密区分
                                              EXCEL_SHEET_COLUMN7,    // ファイルパス
                                              EXCEL_SHEET_COLUMN8,    // 作成者
                                              EXCEL_SHEET_COLUMN9 };  // 最終更新者

                for (int i=0; i < strFormat.Count(); i++)
                {
                    ws.Range[ws.Cells[2, strFormat[i]],
                        ws.Cells[dataGridViewList.RowCount + 1, strFormat[i]]].NumberFormatLocal = "@";
                }
                
                // Excelにデータを貼り付ける
                Excel.Range range = ws.Range[ws.Cells[1, 1], ws.Cells[dataGridViewList.RowCount + 1, EXCEL_OUTPUT_COLUMN_COUNT]];

                range.Value2 = datas;

                // 罫線を付ける
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // 列幅を自動設定
                range.EntireColumn.AutoFit();

                // 20171018 追加（「作成日」と「更新日」の書式を日付形式に設定）
                ws.Range[ws.Cells[2, EXCEL_SHEET_COLUMN3], 
                    ws.Cells[dataGridViewList.RowCount + 1, EXCEL_SHEET_COLUMN4]].NumberFormat = "yyyy/mm/dd";  // step2 iwasa

                // ファイルを保存
                wb.SaveAs(sfd.FileName);

                // ブックを閉じる
                wb.Close();

                // Excel を終了する
                ExcelApp.Quit();

                // COM オブジェクトの参照カウントを解放する (正しくは COM オブジェクトの参照カウントを解放する を参照)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);

                //// キーワードを生成 B秘で出力する
                string keyword = SECRECY_PROPERTY_B + "; " +
                    "" + "; " + settingForm.clsCommonSettting.strOfficeCode + ";";

                //// プロパティ設定を行う
                string error_reason = "";

                DateTime updateTime = new DateTime();

                AccessToProperties atp = new AccessToProperties();
                atp.WriteProperties(sfd.FileName, keyword, ListForm.EXTENSION_EXCEL, false, null,
                                                      ref updateTime, ref error_reason);

                // 完了メッセージ表示
                MessageBox.Show(Resources.msgExportExcelFile, Resources.msgConfirmation, MessageBoxButtons.OK, MessageBoxIcon.Information);

                #endregion
            }
            catch(Exception ex)
            {
                // 20190705 TSE matsuo
                // カレントディレクトリに「Log」フォルダが無ければ作成
                if (!Directory.Exists(EXCLUSION_LOG_PATH + @"\Log"))
                {
                    Directory.CreateDirectory(EXCLUSION_LOG_PATH + @"\Log");
                }

                // エラーログファイル名「errorLog_yyyymmddHHMMss.log」
                using (StreamWriter writer = new StreamWriter(EXCLUSION_LOG_PATH + @"\Log" + @"\FileOutputErr.log", true))    // 末尾に追記
                {
                    writer.WriteLine("");
                    writer.WriteLine("");
                    writer.WriteLine("■■■■■■■■■");
                    writer.WriteLine(DateTime.Now.ToString());
                    writer.WriteLine(ex.ToString());
                    writer.Close();
                }

                // エラーメッセージ表示
                MessageBox.Show(Resources.msgFailedWriteExcelFile, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
            finally
            {
            	// 描画再開
                dataGridViewList.Visible = true;
                dataGridViewList.ResumeLayout();
            }
        }

        #endregion

        #region メソッド

        /// <summary>
        /// リスト表示データ追加処理
        /// </summary>
        /// <param name="strTargefilePath">ファイルリスト</param>
        /// <param name="arrayFileList">格納するアレイリスト</param>
        /// <returns></returns>
        private Boolean SetFileListAdd(string strTargefilePath, ref ArrayList arrayFileList , ref bool fileReadErrFlg)     // 20171012 修正（ファイル読込時にエラーが発生した場合にフラグを更新）
        {
            GridListData clsListData = new GridListData();          // グリッドデータクラス
            Boolean bResult = false;                                // 結果(戻り値)

            try
            {
                // ファイルが存在しない場合は処理終了
                if (!File.Exists(strTargefilePath)) return bResult;

                // ファイル名取得
                clsListData.strFileName = System.IO.Path.GetFileName(strTargefilePath);

                #region<拡張子チェック>
                // 拡張子取得
                int iExtension = ExtensionCheckWithZip(strTargefilePath);

                // 指定拡張子以外の場合はスキップ
                if (iExtension == EXTENSION_NONE) return bResult;

                #endregion

                #region<作成日チェック>

                // 作成日取得
                DateTime dtCreateDate = System.IO.File.GetCreationTime(strTargefilePath);
                clsListData.strCreateDate = dtCreateDate.ToString("yyyy/MM/dd HH:mm:ss");
                #endregion

                #region<更新日チェック>
                // 更新日取得
                DateTime dtUpdateDate = System.IO.File.GetLastWriteTime(strTargefilePath);
                clsListData.strUpdateDate = dtUpdateDate.ToString("yyyy/MM/dd HH:mm:ss");
                #endregion

                string error_reason = "";                               // 20171011 追加（ファイル書込み時にエラーが発生した場合にエラー理由を保存）

                #region<文書分類、機密区分取得>
                switch (iExtension)
                {
                    case EXTENSION_PDF:                             // PDFの場合
                        if (ReadByPDF(strTargefilePath, ref clsListData.strClassNo, ref clsListData.strSecrecyLevel, ref clsListData.strCreator, ref clsListData.strClassNoSerch_hideen, ref clsListData.strClassNo_hideen, ref error_reason) == false)   // 20171228 修正（エラー理由保存）
                        {
                            clsListData.strClassNo = LIST_VIEW_NA;  // 文書分類表示：該当無し
                        }
                        break;
                    case EXTENSION_EXCEL:                           // Excelの場合
                    case EXTENSION_WORD:                            // Wordの場合
                    case EXTENSION_POWERPOINT:                      // PowerPointの場合
                        Boolean bStamp = false;

                        if (ExtensionOpenXMLCheck(strTargefilePath))
                        {
                            // OFFICE2016対応
                            AccessToProperties atp = new AccessToProperties();
                            Dictionary<string, string> propertiesTable = new Dictionary<string, string>();
                            string str_exception = "";
                            propertiesTable = atp.ReadProperties(strTargefilePath, iExtension, ref error_reason, ref str_exception);
                            atp.setProperties(propertiesTable, ref clsListData, ref bStamp, settingForm.clsCommonSettting.strOfficeCode, error_reason);
                        }
                        else
                        {
                            // 20170905 コメントアウト（保存期限、保存年数の追加）
                            // 20171228 追加（エラー理由保存）
                            if (ReadByDSO(strTargefilePath, ref clsListData, ref bStamp, ref error_reason) == false)
                            {
                                clsListData.strClassNo = LIST_VIEW_NA;  // 文書分類表示：該当無し
                            }
                        }
                        break;
                    default:
                        break;
                }
                #endregion

                bool IsZipItem = false;
                string ZipPath = "";

                // Zip内のデータか？
                IsZipItem = CompressedFileChecker.IsCompresstionZipItem(
                    dicCompressedItem,
                    strTargefilePath,
                    out ZipPath
                    );

                // ファイルパス取得
                clsListData.strFilePath = strTargefilePath;

                if (IsZipItem != false)
                {
                    // Zip内のデータの場合
                    clsListData.strFilePath = ZipPath;
                    clsListData.strTmpZipFilePath = strTargefilePath;
                }

                // 20171012 追加（ファイルの事業所名を取得（「no_stamp」「第3パラメータ無」の場合は"HM"））

                // step2 iwasa
                if(IsZipItem)
                {
                    clsListData.strZipFormat = Resources.msgZipFormatContent;
                }
                else if(ExtensionCheckZip(strTargefilePath) == EXTENSION_ZIP)
                {
                    clsListData.strZipFormat = Resources.msgZipFormatFile;
                }
                else
                {
                    clsListData.strZipFormat = Resources.msgZipFormatNot;
                }

                // アレイリストに追加
                arrayFileList.Add(clsListData);

                bResult = true;

                // 20171012 追加（ファイル読み込み時のエラーがあればフラグを更新）
                if (error_reason != "")
                {
                    fileReadErrFlg = true;
                }

            }
            catch (Exception ex)
            {
                // 追加不可能なデータな為スルーする
            }

            return bResult;
        }

        /// <summary>
        /// リスト表示データ絞込追加処理
        /// </summary>
        /// <param name="strTargefilePath">検索データ</param>
        /// <param name="iFileType">ファイルの種類</param>
        /// <param name="strSearchFileName">絞込ファイル名</param>
        /// <param name="dtCreateFrom">作成日From</param>
        /// <param name="dtCreateTo">作成日To</param>
        /// <param name="dtUpdateFrom">更新日From</param>
        /// <param name="dtUpdateTo">更新日To</param>
        /// <returns></returns>
        private Boolean SetFileListRefine(GridListData clsListData, int iFileType, string strSearchFileName, DateTime? dtCreateFrom, DateTime? dtCreateTo, DateTime? dtUpdateFrom, DateTime? dtUpdateTo, Boolean bSecrecyLevelCheck, Boolean bClassNoCheck)
        {
            Boolean bResult = false;                                // 結果(戻り値)

            try
            {
                #region<拡張子チェック>
                // 拡張子取得
                int iExtension = ExtensionCheck(clsListData.strFilePath);
                if (Path.GetExtension(clsListData.strFilePath).Contains("zip") != false)
                {
                    // zipの場合はtemp先を見る
                    iExtension = ExtensionCheck(clsListData.strTmpZipFilePath);
                }

                // 指定拡張子以外の場合はエラー
                if (iExtension == EXTENSION_NONE) return bResult;

                // OFFICEファイル指定でOFFICEファイル以外の場合はエラー
                if ((iFileType == EXTENSION_TYPE_OFFICE) && ((iExtension != EXTENSION_EXCEL) && (iExtension != EXTENSION_WORD) && (iExtension != EXTENSION_POWERPOINT))) return bResult;

                // PDFファイル指定でPDFファイル以外の場合はエラー
                if ((iFileType == EXTENSION_TYPE_PDF) && (iExtension != EXTENSION_PDF)) return bResult;

                #endregion

                #region<ファイル名チェック>
                if (string.IsNullOrEmpty(strSearchFileName) == false)
                {
                    if ((strSearchFileName != "") && (strSearchFileName != "*"))
                    {
                        ClassAttributeSetting clsAs = new ClassAttributeSetting();
                        // 同じ結果を得られる正規表現に変換
                        string regptn = Regex.Escape(strSearchFileName);
                        regptn = clsAs.getSaerchString(regptn);
                        regptn = regptn.Replace(@"\*", ".*?");
                        regptn = regptn.Replace(@"\?", ".");

                        // 一致しない場合はエラー
                        Regex regex = new Regex(regptn);
                        if (!regex.IsMatch(clsAs.getSaerchString(clsListData.strFileName))) return bResult;
                    }
                }
                #endregion

                #region<作成日チェック>
                // 作成日取得
                DateTime dtCreateDate = DateTime.Parse(clsListData.strCreateDate);

                // 作成日範囲外の場合はエラー
                if ((dtCreateFrom != null) && (dtCreateDate < dtCreateFrom)) return bResult;
                if ((dtCreateTo != null) && (dtCreateDate >= dtCreateTo)) return bResult;
                #endregion

                #region<更新日チェック>
                // 更新日取得
                DateTime dtUpdateDate = DateTime.Parse(clsListData.strUpdateDate);

                // <更新日範囲外の場合はエラー
                if ((dtUpdateDate != null) && (dtUpdateDate < dtUpdateFrom)) return bResult;
                if ((dtUpdateTo != null) && (dtUpdateDate >= dtUpdateTo)) return bResult;
                #endregion

                #region<SAB秘チェック>
                // チェックがついていないSAB秘の場合はエラー
                if (bSecrecyLevelCheck)
                {
                    switch (clsListData.strSecrecyLevel)
                    {
                        case SECRECY_PROPERTY_S:
                            if (checkBoxSAB_S.Checked == false) return bResult;
                            break;
                        case SECRECY_PROPERTY_A:
                            if (checkBoxSAB_A.Checked == false) return bResult;
                            break;
                        case SECRECY_PROPERTY_B:
                            if (checkBoxSAB_B.Checked == false) return bResult;
                            break;
                        case "":
                            // 空欄の場合
                            if (checkBoxSAB_None.Checked == false) return bResult;
                            break;
                        default:
                            // S,A,B以外の文字列が入っている場合
                            if (checkBoxSAB_Other.Checked == false) return bResult;
                            break;
                    }
                }
                #endregion

                #region<文書分類チェック>
                if (bClassNoCheck)
                {
                    if ((textBoxClassNo.Text != "") && (textBoxClassNo.Text != "*"))
                    {
                        ClassAttributeSetting clsAs = new ClassAttributeSetting();
                        // 同じ結果を得られる正規表現に変換
                        string regptn = Regex.Escape(textBoxClassNo.Text);
                        regptn = clsAs.getSaerchString(regptn);
                        regptn = regptn.Replace(@"\*", ".*?");
                        regptn = regptn.Replace(@"\?", ".");

                        // 一致しない場合はエラー
                        Regex regex = new Regex(regptn);
                        //// 20170928 TSE kitada 検索対象を機密区分から機密区分_hiddenに変更
                        // 20171013 修正 （検索対象を機密区分_hiddenから文書分類検索用_hiddenに変更）
                        if (!regex.IsMatch(clsAs.getSaerchString(clsListData.strClassNoSerch_hideen))) return bResult;
                    }
                }
                #endregion

                bResult = true;
            }
            catch
            {
                // 追加不可能なデータな為スルーする
            }

            return bResult;
        }

        /// <summary>
        /// 画面初期化処理
        /// </summary>
        private void Initialize()
        {
            // フォルダー指定
            textBoxFolderPath.Text = "";

            // サブフォルダ―以下も含む
            radioButtonSubFolderInc.Checked = true;

            // All指定
            //radioButtonTypeOfficePdf.Checked = true;
            radioButtonTypeAll.Checked = true;

            // ファイル名指定
            textBoxFileName.Text = "";

            // 作成年月日
            checkBoxCreate.Checked = false;
            dateTimePickerCreateFrom.Value = DateTime.Today.AddYears(-1);
            dateTimePickerCreateTo.Value = DateTime.Today;

            // 更新年月日
            checkBoxUpdate.Checked = false;
            dateTimePickerUpdateFrom.Value = DateTime.Today.AddYears(-1);
            dateTimePickerUpdateTo.Value = DateTime.Today;

            // S/A/B秘指定
            checkBoxSAB_All.Checked = true;

            // 文書分類指定
            textBoxClassNo.Text = "";

            // 全てを同じ機密区分・文書分類に一括変更
            radioButtonBatch.Checked = true;

            // リストの各列の入力設定
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.ZipFormat].ReadOnly = true;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.FileName].ReadOnly = true;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.FileType].ReadOnly = true;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.CreatedDate].ReadOnly = true;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.UpdatedDate].ReadOnly = true;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.Select].ReadOnly = false;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.DocumentType].ReadOnly = true;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.SecretType].ReadOnly = true;
            this.dataGridViewList.Columns[(int)DataGridViewColumnIndex.FilePath].ReadOnly = true;

            // ステータスバー
            labelStatus.Text = "";

            // リストのファイルタイプ無し
            IListFileType = EXTENSION_TYPE_NONE;


            System.Diagnostics.FileVersionInfo ver =
            System.Diagnostics.FileVersionInfo.GetVersionInfo(
            System.Reflection.Assembly.GetExecutingAssembly().Location);

            // タイトル
            string AssemblyName = ver.FileVersion;
            this.Text = this.Text + " " + AssemblyName;

            this.Refresh();
        }

        /// <summary>
        /// 実行中の画面コントロール制御
        /// </summary>
        /// <param name="bEnabled"></param>
        private void ProcControlEnabled(Boolean bEnabled)
        {
            // ストップフラグOFF
            bStopFlg = false;

            // ステータス初期化
            labelStatus.Text = "";

            // メッセージボックス
            // 20190531 TSE kitada 画面リサイズ対応でパネルを画面中央に移動
            panelProcess.Top = (this.Height - panelProcess.Height) / 2;
            panelProcess.Left = (this.Width - panelProcess.Width) / 2;

            panelProcess.Visible = !bEnabled;
            
            // 入力条件パネル
            panelInput.Enabled = bEnabled;

            // リスト
            dataGridViewList.Enabled = bEnabled;

            // 出力条件パネル
            panelOutput.Enabled = bEnabled;

            // 設定開始ボタン
            buttonStart.Enabled = bEnabled;

            // Excel出力ボタン
            buttonExcelOutput.Enabled = bEnabled;
        }

        /// <summary>
        /// ディレイ
        /// </summary>
        /// <param name="dTimerCnt"></param>
        private void Delay(double dTimerCnt)
        {
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            while ((double)sw.ElapsedTicks / System.Diagnostics.Stopwatch.Frequency < dTimerCnt) { System.Windows.Forms.Application.DoEvents(); }
            sw.Stop();
        }

        /// <summary>
        /// Officeプロパティ書き込み
        /// 20171213 TSE demachi CHG グリッドに表示する更新日を取得するため引数にref DateTime LastWriteTimeを追加
        /// </summary>
        /// <param name="file"></param>
        /// <returns>true:成功 false:失敗</returns>
        public Boolean WriteByDSO(string strFile, string strKeyword, Boolean bFileInfoUpdate, ref DateTime LastWriteTime, ref string error_reason)  // 20171011 修正（ファイル書込み時のエラー理由を保存）
        {
            Boolean bResult = false;
            LastWriteTime = new DateTime();
            DSOFile.OleDocumentProperties ducProperty = new DSOFile.OleDocumentProperties();
            DSOFile.SummaryProperties summary;

            try
            {
                // 20171011 修正（ファイル書込み時のエラー理由を保存）
                // 読み取り専用チェック 読み取り専用である場合はエラーとする
                FileAttributes fas = File.GetAttributes(strFile);
                if ((fas & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    // 20171228 追加（エラー理由保存）
                    error_reason = ListForm.LIST_VIEW_NA;
                    return bResult;
                }

                try
                {
                    // ファイル情報読み取り
                    LastWriteTime = System.IO.File.GetLastWriteTime(strFile);
                    // ファイルを開く
                    ducProperty.Open(strFile, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);

                    // プロパティの取得
                    summary = ducProperty.SummaryProperties;

                    // キーワードに書き込み
                    summary.Keywords = strKeyword;

                    // プロパティはクリアする
                    summary.Category = "";

                    // dsofileクローズ
                    ducProperty.Close(true);

                    // 更新日は更新しない場合は元の更新日を上書きする
                    if (!bFileInfoUpdate)
                    {
                        System.IO.File.SetLastWriteTime(strFile, LastWriteTime);
                    }

                    bResult = true;
                }
                catch
                {
                    // 20171011 修正（ファイル書込み時のエラー理由を保存）
                    // ファイルが開かれている場合のエラー
                    error_reason = LIST_VIEW_NA;
                }
            }
            catch(Exception ex)
            {
                // 20171011 修正（ファイル書込み時のエラー理由を保存）
                // ファイルが存在しない場合のエラー
                error_reason = LIST_VIEW_NA;
            }

            return bResult;
        }

        /// <summary>
        /// DSOによるOfficeの読み込み
        /// </summary>
        /// <param name="strFile">ファイル</param>
        /// <param name="clsListData">Gridへ書き込むためのリスト</param>
        /// <returns>true:成功 false:失敗</returns>
        public Boolean ReadByDSO(string strFile, ref ListForm.GridListData clsListData, ref bool bStamp ,ref string error_reason)
        {
            Boolean bResult = false;
            DSOFile.OleDocumentProperties ducProperty = new DSOFile.OleDocumentProperties();
            DSOFile.SummaryProperties summary;

            try
            {
                // ファイルを開く
                ducProperty.Open(strFile, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);

                // プロパティの取得
                summary = ducProperty.SummaryProperties;

                // カテゴリーから読み込み
                clsListData.strClassNo = (summary.Category != null) ? summary.Category : "";

                // タグ情報があった場合は機密区分、文書分類番号、スタンプの有無をセットする
                if (summary.Keywords != null)
                {
                    string[] strPropertyData = summary.Keywords.Split(';');

                    // 機密区分、文書分類番号、スタンプの有無の3項目以上ある場合
                    if (strPropertyData.Count() >= 3)
                    {
                        if (strPropertyData[2].Trim() == NO_STAMP) bStamp = false;
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
                        // 20171006 修正（自事業所コードとファイル事業所コードによる分岐を追加）
                        else if (strPropertyData[2].Trim() != settingForm.clsCommonSettting.strOfficeCode)
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

                    // 20170915追加（作成者情報の取得を追加）
                    clsListData.strCreator = summary.Author;

                    // 20170915追加（最終更新者情報の取得を追加）
                    clsListData.strLastModifiedBy = summary.LastSavedBy;

                    // 20171013 追加（文書分類検索用の隠しカラムに文書分類をそのままコピー）
                    clsListData.strClassNoSerch_hideen = clsListData.strClassNo;
                }
                bResult = true;
            }
            catch
            {
                // 20171228 追加（エラー理由保存）
                error_reason = ListForm.LIST_VIEW_NA;
            }
            finally
            {
                ducProperty.Close(true);
            }
            return bResult;
        }

        /// <summary>
        /// PDFプロパティ書き込み
        /// 20171213 TSE demachi CHG グリッドに表示する更新日を取得するため引数にref DateTime LastWriteTimeを追加
        /// </summary>
        /// <param name="file"></param>
        /// <returns>true:成功 false:失敗</returns>
        public Boolean WriteByPDF(string strFilePath, string strKeyword, Boolean bFileInfoUpdate, ref DateTime LastWriteTime, ref string error_reason)  // 20171011 修正（ファイル書込み時のエラー理由を保存）
        {
            string strTempFilePath;                                 // PDFファイル一時書き込みファイルパス
            Boolean bResult = false;

            try
            {
                // PDFファイル一時書き込みファイルパス作成
                int iFileNo = 0;
                do
                {
                    strTempFilePath = string.Format(PDF_TEMPFILEPATH_FORMAT, Path.GetTempPath(), iFileNo++);
                } while (System.IO.File.Exists(strTempFilePath) == true);

                // ファイル情報読み取り
                LastWriteTime = System.IO.File.GetLastWriteTime(strFilePath);

                // 20171011 修正（ファイル書込み時のエラー理由を保存）
                try
                {
                    // PDFReader作成
                    PdfReader pdfReader = new PdfReader(strFilePath);

                    try
                    {
                        // PDFファイル読み込み FileStream
                        using (FileStream fs = new FileStream(strTempFilePath, FileMode.Create, FileAccess.Write))
                        {
                            // PDFファイル読み込み PdfStamper
                            using (PdfStamper pdfStamper = new PdfStamper(pdfReader, fs))
                            {
                                // プロパティ書き込み
                                Dictionary<String, String> info = new Dictionary<String, String>();
                                info["Keywords"] = strKeyword;
                                pdfStamper.MoreInfo = info;

                                pdfStamper.Close();
                                pdfReader.Close();
                            }
                        }

                        // ファイルコピー
                        File.Copy(strTempFilePath, strFilePath, true);

                        // テンポラリファイル削除
                        File.Delete(strTempFilePath);


                        // 更新日は更新しない場合は元の更新日を上書きする
                        if (!bFileInfoUpdate)
                        {
                            System.IO.File.SetLastWriteTime(strFilePath, LastWriteTime);
                        }

                        bResult = true;
                    }
                    catch
                    {
                        // ファイルが開かれている場合のエラー
                        error_reason = LIST_VIEW_NA;
                    }
                }
                catch
                {
                    // ファイルが存在しない場合のエラー
                    error_reason = LIST_VIEW_NA;
                }                
            }
            catch(Exception e)
            {
                // 読取専用、ファイ存在しない以外のエラー
                error_reason = LIST_VIEW_NA;
            }
            return bResult;
        }

        /// <summary>
        /// PDFプロパティ読み込み
        /// </summary>
        /// <param name="file"></param>
        /// <returns>true:成功 false:失敗</returns>
        public Boolean ReadByPDF(string strFile, ref string strClassNo, ref string strSecrecyLevel, ref string strCreator, ref string strClassNoSerch_hideen , ref string strClassNo_hideen, ref string error_reason) // 20171228修正（エラー理由保存）
        {
            Boolean bResult = false;

            try
            {
                // PDFの情報取得
                PdfReader reader = new PdfReader(strFile);

                // プロパティ情報をリスト化
                List<string> name = new List<string>(reader.Info.Keys);
                List<string> val = new List<string>(reader.Info.Values);

                // PDFファイルクローズ
                reader.Close();

                // KeyWordsのインデックスを取得
                int Property_Count = name.IndexOf("Keywords");

                // プロパティ情報がない場合は処理を抜ける
                if (!name.Contains("Keywords")) return true;

                // プロパティ読み込み
                string strBuf = val[Property_Count];
                string[] stBuf = strBuf.Split(';');

                // 「;」区切りで2項目未満の場合は処理を抜ける
                if (stBuf.Count() >= 2)
                {
                    // 機密区分設定
                    switch (stBuf[0])
                    {
                        case SECRECY_PROPERTY_S:
                            strSecrecyLevel = SECRECY_PROPERTY_S;
                            break;
                        case SECRECY_PROPERTY_A:
                            strSecrecyLevel = SECRECY_PROPERTY_A;
                            break;
                        case SECRECY_PROPERTY_B:
                            strSecrecyLevel = SECRECY_PROPERTY_B;
                            break;
                        default:
                            strSecrecyLevel = "";
                            break;
                    }

                    // パラメータが2つで
                    if (stBuf.Count() == 2)
                    {
                        strClassNo = stBuf[1].Trim();
                    }
                    // パラメータの3つ目が「no_stamp」or「無」で
                    else if (stBuf[2].Trim() == NO_STAMP || stBuf[2].Trim() == "")
                    {
                        strClassNo = stBuf[1].Trim();
                    }
                    // パラメータの3つ目が自事業所でない場合
                    else if (stBuf[2].Trim() != settingForm.clsCommonSettting.strOfficeCode)
                    {
                        strClassNo = stBuf[2].Trim();
                    }
                    else
                    {
                        strClassNo = stBuf[1].Trim();
                    }

                    // 20171228 修正（作成者情報がない場合にアプリが落ちないように修正）
                    if (name.IndexOf("Author") != -1)
                    {
                        strCreator = val[name.IndexOf("Author")];
                    }                  
                   
                    // 20171013 追加（文書分類検索用の隠しカラムに文書分類をそのままコピー）
                    strClassNoSerch_hideen = strClassNo;

                    // 20171018 追加（文書分類番号用の隠しカラムに文書分類番号を保存）
                    if (stBuf.Count() >= 2)
                    {
                        // 文書分類番号用の隠しカラムに保存
                        strClassNo_hideen = stBuf[1].Trim();
                    }
                }
                bResult = true;
            }
            catch
            {
                // 20171228 追加（エラー理由保存）
                error_reason = ListForm.LIST_VIEW_NA;
            }

            return bResult;
        }

        /// <summary>
        /// 拡張子種別チェック
        /// </summary>
        /// <param name="strFilePath">チェックするファイルフルパス</param>
        /// <returns>拡張子種別</returns>
        public int ExtensionCheck(string strFilePath)
        {
            // 拡張子初期設定無し
            int iExtension = EXTENSION_NONE;

            // 拡張子取得
            string strTargetExtension = System.IO.Path.GetExtension(strFilePath);
            for (int i = 0; i < StrExtension_Narrow.Count(); i++)
            {
                if (StrExtension_Narrow[i] == strTargetExtension.ToLower())
                {
                    iExtension = iExtension_Narrow[i];
                    break;
                }
            }

            return iExtension;
        }

        /// <summary>
        /// 拡張子種別チェック(OpenXMLでの書き込み対象かを判定)
        /// </summary>
        /// <param name="strFilePath">チェックするファイルフルパス</param>
        /// <returns>拡張子種別</returns>
        public bool ExtensionOpenXMLCheck(string strFilePath)
        {
            bool bRet = false;

            // 拡張子取得
            string strTargetExtension = System.IO.Path.GetExtension(strFilePath);

            foreach(string narrow in StrOpenXML_Narrow)
            {
                if(strTargetExtension == narrow)
                {
                    // OpenXMLの対象
                    bRet = true;
                    break;
                }
            }

            return bRet;
        }

        /// <summary>
        /// データグリッドビューの表示更新
        /// </summary>
        /// <param name="arrayFileList"></param>
        private void dataGridView_redraw(ArrayList arrayFileList)
        {
            // データグリッドの内容をクリア
            dataGridViewList.Rows.Clear();
            
            checkBoxAll.Checked = false;

            // ソート状態をクリア
            ClearSortDirection();

            // step2 iwasa
            List<string> tempList = new List<string>();

            CurrentViewZipCount = 0;

            // リストにデータ追加
            foreach (GridListData clsListData in arrayFileList)
            {
                // 20171009 追加（他事業所ファイルの場合は事業所名を表示）
                // 20171009 追加（他事業所ファイルの場合は文書分類に事業所名を表示）
                string ClassNoOrOfficeName;
                if (settingForm.clsCommonSettting.strOfficeCode != clsListData.strClassNo)
                {
                    ClassNoOrOfficeName = clsListData.strClassNo;
                }
                else
                {
                    ClassNoOrOfficeName = clsListData.strClassNo_hideen;
                }

                // 20171212 TSE demachi ADD グリッドに表示する変数がNullの場合空白("")に置き換える
                GridListDataNullToEmpty(clsListData);

                // 20171213 TSE demachi CHG 万が一clsListData.strFilePathが空白の場合Path.GetDirectoryNameでエラーになるため、回避処理を追加
                if (clsListData.strFilePath.Length == 0)
                {
                    dataGridViewList.Rows.Add(clsListData.strZipFormat, clsListData.strFileName, clsListData.strFileType, clsListData.strCreateDate,    // step2 iwasa
                    clsListData.strUpdateDate, false, ClassNoOrOfficeName, clsListData.strSecrecyLevel,
                    clsListData.strFilePath, "", "",
                    clsListData.strCreator, clsListData.strLastModifiedBy, clsListData.strClassNo_hideen, "", (0).ToString());
                }
                else if (clsListData.strTmpZipFilePath.Length != 0)
                {
                    // step2 iwasa
                    // zipの一覧表示の変更
                    if (tempList.Contains(clsListData.strFilePath) == false)
                    {
                        CurrentViewZipCount++;

                        // zipファイル名は一つだけ表示する
                        tempList.Add(clsListData.strFilePath);

                        string CreateTime = System.IO.File.GetCreationTime(clsListData.strFilePath).ToString("yyyy/MM/dd HH:mm:ss");
                        string UpdateTime = System.IO.File.GetCreationTime(clsListData.strFilePath).ToString("yyyy/MM/dd HH:mm:ss");

                        dataGridViewList.Rows.Add(Resources.msgZipFormatFile, Path.GetFileName(clsListData.strFilePath), clsListData.strFileType,
                        CreateTime, UpdateTime, false, "", "",
                        clsListData.strFilePath, "", "",
                        clsListData.strCreator, clsListData.strLastModifiedBy, clsListData.strClassNo_hideen,
                        clsListData.strFilePath, CurrentViewZipCount.ToString()
                        );

                        // ZIPファイルは選択できないようにする
                        dataGridViewList[(int)DataGridViewColumnIndex.Select, dataGridViewList.RowCount - 1].ReadOnly = true;
                    }

                    // 20200805 ZIPファイルの中身の場合
                    dataGridViewList.Rows.Add(clsListData.strZipFormat, clsListData.strFileName, clsListData.strFileType, clsListData.strCreateDate,
                    clsListData.strUpdateDate, false, ClassNoOrOfficeName, clsListData.strSecrecyLevel,
                    clsListData.strFilePath, "", "",
                    clsListData.strCreator, clsListData.strLastModifiedBy, clsListData.strClassNo_hideen,
                    clsListData.strTmpZipFilePath, CurrentViewZipCount.ToString()
                    );
                    
                }
                else {
                    dataGridViewList.Rows.Add(clsListData.strZipFormat, clsListData.strFileName, clsListData.strFileType, clsListData.strCreateDate,
                    clsListData.strUpdateDate, false, ClassNoOrOfficeName, clsListData.strSecrecyLevel,
                    Path.GetDirectoryName(clsListData.strFilePath), "", "",
                    clsListData.strCreator, clsListData.strLastModifiedBy, clsListData.strClassNo_hideen, "", (0).ToString().ToString());
                }
                
            }

            if (IListFileType == EXTENSION_TYPE_ALL)
            {
                // パスワード付きzipを一覧に表示する step2 iwasa
                foreach (KeyValuePair<string, List<string>> pair in dicPasswordZip)
                {
                    List<string> file = pair.Value;

                    CurrentViewZipCount++;
                    string CreateTime = System.IO.File.GetCreationTime(pair.Key).ToString("yyyy/MM/dd HH:mm:ss");
                    string UpdateTime = System.IO.File.GetCreationTime(pair.Key).ToString("yyyy/MM/dd HH:mm:ss");

                    string msgZip = Resources.msgZipDecompress + "(" + Resources.msgZipPasswordError + ")";

                    // zipのファイル名を表示
                    dataGridViewList.Rows.Add(Resources.msgZipFormatFile, Path.GetFileName(pair.Key), "", CreateTime,
                    UpdateTime, false, msgZip, "",
                    Path.GetDirectoryName(pair.Key), "", "",
                    "", "", "", "", CurrentViewZipCount.ToString());

                    // 解凍不可ZIPファイルは選択できないようにする
                    dataGridViewList[(int)DataGridViewColumnIndex.Select, dataGridViewList.RowCount - 1].ReadOnly = true;

                    // zipの中身を表示
                    foreach (string item in file)
                    {
                        if (string.IsNullOrEmpty(item) == false)
                        {
                            dataGridViewList.Rows.Add(Resources.msgZipFormatContent, item, "", "-",
                            "-", false, "", msgZip,
                            "", "", "",
                            "", "", "", "", CurrentViewZipCount.ToString());

                            // 解凍不可なので選択できないようにする
                            dataGridViewList[(int)DataGridViewColumnIndex.Select, dataGridViewList.RowCount - 1].ReadOnly = true;
                        }
                    }
                }

                // エラーで解凍不可のzipを一覧表示する
                foreach (KeyValuePair<string, List<string>> pair in dicErrorZip)
                {
                    List<string> file = pair.Value;

                    CurrentViewZipCount++;
                    string CreateTime = System.IO.File.GetCreationTime(pair.Key).ToString("yyyy/MM/dd HH:mm:ss");
                    string UpdateTime = System.IO.File.GetCreationTime(pair.Key).ToString("yyyy/MM/dd HH:mm:ss");

                    string msgZip = Resources.msgZipDecompress + "(" + Resources.msgZipFileError + ")";

                    // zipのファイル名を表示
                    dataGridViewList.Rows.Add(Resources.msgZipFormatFile, Path.GetFileName(pair.Key), "", CreateTime,
                    UpdateTime, false, msgZip, "",
                    Path.GetDirectoryName(pair.Key), "", "",
                    "", "", "", "", CurrentViewZipCount.ToString());

                    // 解凍不可ZIPファイルは選択できないようにする
                    dataGridViewList[(int)DataGridViewColumnIndex.Select, dataGridViewList.RowCount - 1].ReadOnly = true;

                    // zipの中身を表示
                    foreach (string item in file)
                    {
                        dataGridViewList.Rows.Add(Resources.msgZipFormatContent, item, "", "-",
                        "-", false, "", msgZip,
                        "", "", "",
                        "", "", "", "", CurrentViewZipCount.ToString());

                        // 解凍不可なので選択できないようにする
                        dataGridViewList[(int)DataGridViewColumnIndex.Select, dataGridViewList.RowCount - 1].ReadOnly = true;
                    }
                }
            }
        }


        /// <summary>
        /// 20171212 TSE demachi ADD グリッドに表示する変数がNullの場合空白("")に置き換える
        /// </summary>
        /// <param name="clsListData"></param>
        private void GridListDataNullToEmpty(GridListData clsListData)
        {
            // Nullの場合空白に置き換える
            if (clsListData.strZipFormat == null) clsListData.strZipFormat = "";    // step2 iwasa
            if (clsListData.strFileName == null) clsListData.strFileName = "";
            if (clsListData.strFileType == null) clsListData.strFileType = "";
            if (clsListData.strCreateDate == null) clsListData.strCreateDate = "";
            if (clsListData.strUpdateDate == null) clsListData.strUpdateDate = "";
            if (clsListData.strSecrecyLevel == null) clsListData.strSecrecyLevel = "";
            if (clsListData.strFilePath == null) clsListData.strFilePath = "";
            if (clsListData.strCreator == null) clsListData.strCreator = "";
            if (clsListData.strLastModifiedBy == null) clsListData.strLastModifiedBy = "";
            if (clsListData.strClassNo_hideen == null) clsListData.strClassNo_hideen = "";
            if (clsListData.strClassNo == null) clsListData.strClassNo = "";
            if (clsListData.strClassNoSerch_hideen == null) clsListData.strClassNoSerch_hideen = "";
        }
        #endregion

        #region メニュー

        /// <summary>
        /// 20170904 追加（グリッド内で右クリックからファイルを開く）
        /// </summary>
        private void ファイルを開くToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridViewList.Rows.Count > 0)
            {
                foreach (DataGridViewRow r in dataGridViewList.SelectedRows)
                {
                    if (!r.IsNewRow)
                    {
                        // リストで選択されている行のファイルパスを取得
                        string IsZipFormat = r.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString();
                        if (string.IsNullOrEmpty(IsZipFormat) == false)
                        {
                            // ZIP本体またはZIPファイルの中身の場合
                            MessageBox.Show(Resources.msgErrorFileOpen, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                            return;
                        }


                        // リストで選択されている行のファイルパスを取得
                        string open_file_pass = r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString() + @"\" + r.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString();

                        if (string.IsNullOrEmpty(r.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString()) == false)
                        {
                            // zipファイルの中身の場合
                            open_file_pass = r.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString();
                        }

                        // ファイルの拡張子を確認
                        string extension = Path.GetExtension(open_file_pass);

                        // 対象のファイルに関連付けられたコマンドを取得（開くアプリケーションを取得）
                        string command = "";

                        if (extension.Contains("pdf") == false)
                        {
                            command = FindAssociatedCommand(open_file_pass, "open");
                        }
                        else
                        {
                            command = FindPdfAssociatedCommand(open_file_pass, "open");
                        }

                        // .exeの位置
                        int exe_pos = 0;
                        if (command.IndexOf(".EXE") > 0)
                        {
                            exe_pos = command.IndexOf(".EXE");
                        }
                        else if (command.IndexOf(".exe") > 0)
                        {
                            exe_pos = command.IndexOf(".exe");
                        }

                        // 先頭の"/"と .exeの後ろの文字列を取り除く
                        command = command.Substring(1, exe_pos + 3);

                        // ファイルパスに空白が含まれることを考慮して前後を"でくくる
                        open_file_pass = "\"" + open_file_pass + "\"";

                        // 関連付けられたアプリケーションでファイルを開く
                        System.Diagnostics.Process.Start(command, open_file_pass);
                    }
                }
            }
        }

        /// <summary>
        /// 指定されたファイルに関連付けられたコマンドを取得する
        /// </summary>
        /// <param name="fileName">関連付けを調べるファイル</param>
        /// <param name="extra">アクション(open,print,editなど)</param>
        /// <returns>取得できた時は、コマンド(実行ファイルのパス+コマンドライン引数)。
        /// 取得できなかった時は、空の文字列。</returns>
        /// <example>
        /// "1.txt"ファイルの"open"に関連付けられたコマンドを取得する例
        /// <code>
        /// string command = FindAssociatedCommand("1.txt", "open");
        /// </code>
        /// </example>
        public static string FindAssociatedCommand(
            string fileName, string extra)
        {
            //拡張子を取得
            string extName = System.IO.Path.GetExtension(fileName);
            if (extName.Length == 0 || extName[0] != '.')
            {
                return string.Empty;
            }

            //HKEY_CLASSES_ROOT\(extName)\shell があれば、
            //HKEY_CLASSES_ROOT\(extName)\shell\(extra)\command の標準値を返す
            if (ExistClassesRootKey(extName + @"\shell"))
            {
                return GetShellCommandFromClassesRoot(extName, extra);
            }

            //HKEY_CLASSES_ROOT\(extName) の標準値を取得する
            string fileType = GetDefaultValueFromClassesRoot(extName);
            if (fileType.Length == 0)
            {
                return string.Empty;
            }

            //HKEY_CLASSES_ROOT\(fileType)\shell\(extra)\command の標準値を返す
            return GetShellCommandFromClassesRoot(fileType, extra);
        }

        /// <summary>
        /// 指定されたファイルに関連付けられたコマンドを取得する(PDF)
        /// </summary>
        /// <param name="fileName">関連付けを調べるファイル</param>
        /// <param name="extra">アクション(open,print,editなど)</param>
        public static string FindPdfAssociatedCommand(
            string fileName, string extra)
        {
            //拡張子を取得
            string extName = System.IO.Path.GetExtension(fileName);
            if (extName.Length == 0 || extName[0] != '.')
            {
                return string.Empty;
            }

            //HKEY_CLASSES_ROOT\(extName) の標準値を取得する
            string fileType = GetDefaultValueFromClassesRoot(extName);
            if (fileType.Length == 0)
            {
                return string.Empty;
            }

            //HKEY_CLASSES_ROOT\(fileType)\shell\(extra)\command の標準値を返す
            return GetShellCommandFromClassesRoot(fileType, extra);
        }

        /// <summary>
        /// 指定したレジストリのオブジェクトを取得
        /// </summary>
        /// <param name="keyName">キー</param>
        private static bool ExistClassesRootKey(string keyName)
        {
            Microsoft.Win32.RegistryKey regKey =
                Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(keyName);
            if (regKey == null)
            {
                return false;
            }
            regKey.Close();
            return true;
        }

        /// <summary>
        /// 指定したレジストリのオブジェクトを取得
        /// </summary>
        /// <param name="keyName">キー</param>
        private static string GetDefaultValueFromClassesRoot(string keyName)
        {
            Microsoft.Win32.RegistryKey regKey =
                Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(keyName);
            if (regKey == null)
            {
                return string.Empty;
            }
            string val = (string)regKey.GetValue(string.Empty, string.Empty);
            regKey.Close();

            return val;
        }

        /// <summary>
        /// 指定したレジストリのオブジェクトを取得
        /// </summary>
        /// <param name="fileName">関連付けを調べるファイル</param>
        /// <param name="extra">アクション(open,print,editなど)</param>
        private static string GetShellCommandFromClassesRoot(
            string fileType, string extra)
        {
            if (extra.Length == 0)
            {
                //アクションが指定されていない時は、既定のアクションを取得する
                extra = GetDefaultValueFromClassesRoot(fileType + @"shell")
                    .Split(',')[0];
                if (extra.Length == 0)
                {
                    extra = "open";
                }
            }
            return GetDefaultValueFromClassesRoot(
                string.Format(@"{0}\shell\{1}\command", fileType, extra));
        }

        /// <summary>
        /// 20170904 追加（グリッド内で右クリックからリストから削除）
        /// </summary>
        private void 削除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridViewList.Rows.Count > 0)
            {
                //リストで選択されている行を削除する
                foreach (DataGridViewRow r in dataGridViewList.SelectedRows)
                {
                    if (!r.IsNewRow)
                    {
                        // リストで選択されている行のファイルパスを取得
                        string IsZipFormat = r.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString();
                        if (string.IsNullOrEmpty(IsZipFormat) == false)
                        {
                            DialogResult dr = MessageBox.Show(
                                Resources.msgZipListDelete,
                                Resources.msgConfirmation,
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Exclamation);

                            // ZIP本体またはZIPファイルの中身の場合
                            if (dr == System.Windows.Forms.DialogResult.OK)
                            {
                                // はい

                                // zipとそれに入っているファイルをリストからすべて削除

                                string CommonZipPath = r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString();

                                List<DataGridViewRow> listRow = new List<DataGridViewRow>();
                                foreach (DataGridViewRow targetRow in dataGridViewList.Rows)
                                {
                                    if (targetRow.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value != null)
                                    {
                                        string targetPath = targetRow.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString();
                                        if (targetPath == CommonZipPath)
                                        {
                                            // 同一ZIPデータ
                                            listRow.Add(targetRow);
                                        }
                                    }
                                }

                                foreach (DataGridViewRow delRow in listRow)
                                {
                                    // ハッシュテーブルから削除
                                    htList.Remove(delRow.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString());

                                    // 内部保持用ファイルリストから削除
                                    for (int i = 0; i < ArrayFileList.Count; i++)
                                    {
                                        GridListData clsListData = (GridListData)ArrayFileList[i];
                                        if (clsListData.strTmpZipFilePath == delRow.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString())
                                        {
                                            ArrayFileList.RemoveAt(i);
                                            break;
                                        }
                                    }

                                    // リストから削除
                                    dataGridViewList.Rows.Remove(delRow);
                                }
                            }

                            // キャンセル
                            return;
                        }

                        // ハッシュテーブルから削除
                        htList.Remove(r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString());

                        // 内部保持用ファイルリストから削除
                        for (int i = 0; i < ArrayFileList.Count; i++)
                        {
                            GridListData clsListData = (GridListData)ArrayFileList[i];
                            if (clsListData.strFilePath == r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString())
                            {
                                ArrayFileList.RemoveAt(i);
                                break;
                            }
                        }

                        // リストから削除
                        dataGridViewList.Rows.Remove(r);
                    }
                }
            }
        }

        /// <summary>
        /// 右クリック→フォルダを開く
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStrip_OpenFolder_Click(object sender, EventArgs e)
        {
            if (dataGridViewList.Rows.Count > 0)
            {
                foreach (DataGridViewRow r in dataGridViewList.SelectedRows)
                {
                    if (!r.IsNewRow)
                    {
                        // リストで選択されている行のファイルパスを取得
                        string IsZipFormat = r.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString();
                        if (string.IsNullOrEmpty(IsZipFormat) == false)
                        {
                            // ZIP本体またはZIPファイルの中身の場合
                            MessageBox.Show(Resources.msgErrorFolderOpen, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                            return;
                        }

                        // リストで選択されている行のファイルパスを取得
                        string open_folder_pass = r.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString();

                        // フォルダを開く
                        System.Diagnostics.Process.Start(open_folder_pass);
                    }
                }
            }
        }

        /// <summary>
        /// 検索文字列内の「,」をDELIMITERに置換する関数
        /// </summary>
        /// <param name="serch_str">検索文字列</param>
        /// <param name="start_str">検索開始番号</param>
        /// <param name="end_str">検索終了番号</param>
        /// <param name="serch_index">検索インデックス</param>
        public void ReplaceDelimiter(ref string serch_str, int start_str, int end_str, ref int serch_index)
        {
            // 1つ目の「"」と2つ目の「"」の間の1データを取得
            string data = serch_str.Substring(start_str + 1, end_str - start_str - 1);
            // 1つ目の「"」と2つ目の「"」の間の「,」を■DELIMITER■に置換
            string delimiter_data = data.Replace(",", ClassAttributeSetting.DELIMITER);
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

        /// <summary>
        /// 拡張子種別チェック(zip含む) step2 iwasa
        /// </summary>
        /// <param name="strFilePath">チェックするファイルフルパス</param>
        /// <returns>拡張子種別</returns>
        public int ExtensionCheckWithZip(string strFilePath)
        {
            // 拡張子初期設定無し
            int iExtension = EXTENSION_NONE;

            List<string> lstStrWithZip = new List<string>();
            lstStrWithZip = StrExtension_Narrow.ToList();

            List<int> lstIntWithZip = new List<int>();
            lstIntWithZip = iExtension_Narrow.ToList();

            if (checkBoxZipTarget.Checked == true)
            {
                lstStrWithZip.Add(".zip");
                lstIntWithZip.Add(EXTENSION_ZIP);
            }
            
            // 拡張子取得
            string strTargetExtension = System.IO.Path.GetExtension(strFilePath);
            for (int i = 0; i < lstStrWithZip.Count(); i++)
            {
                if (lstStrWithZip[i] == strTargetExtension.ToLower())
                {
                    iExtension = lstIntWithZip[i];
                    break;
                }
            }

            return iExtension;
        }

        /// <summary>
        /// zipチェック
        /// </summary>
        /// <param name="strFilePath">チェックするファイルフルパス</param>
        /// <returns>拡張子種別</returns>
        public int ExtensionCheckZip(string strFilePath)
        {
            // 拡張子初期設定無し
            int iExtension = EXTENSION_NONE;

            // 拡張子取得
            string strTargetExtension = System.IO.Path.GetExtension(strFilePath);

            if (strTargetExtension.ToLower() == ".zip")
            {
                iExtension = EXTENSION_ZIP;
            }
            return iExtension;
        }

#endregion


        #region グリッドソート

        /// <summary>
        /// データグリッドビューヘッダクリックイベント
        /// </summary>
        private void dataGridViewList_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                switch (e.ColumnIndex)
                {
                    case (int)DataGridViewColumnIndex.ZipFormat:
                        // ZIPクリック
                        SortedZipFormat(e);
                        break;
                    case (int)DataGridViewColumnIndex.FileName:
                        // ファイル名クリック
                        SortedFileName(e);
                        break;
                    case (int)DataGridViewColumnIndex.CreatedDate:
                        // 作成日クリック
                        SortedCreateTime(e);
                        break;
                    case (int)DataGridViewColumnIndex.UpdatedDate:
                        // 更新日クリック
                        SortedUpdateTime(e);
                        break;
                    case (int)DataGridViewColumnIndex.Select:
                        // 選択クリック
                        break;
                    case (int)DataGridViewColumnIndex.DocumentType:
                        // 文書分類クリック
                        SortedClassNo(e);
                        break;
                    case (int)DataGridViewColumnIndex.SecretType:
                        // 機密区分クリック
                        SortedSecrecyLevel(e);
                        break;
                    case (int)DataGridViewColumnIndex.FilePath:
                        // ファイルパスクリック
                        SortedFilePath(e);
                        break;
                    default:
                        // 何もしない
                        break;
                }
            }
            catch(Exception ex)
            {
                // 原因不明のエラー
                MessageBox.Show(ex.ToString(),
                     Resources.msgError,
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// ソート状態をクリアする
        /// </summary>
        private void ClearSortDirection()
        {
            ZipSortDirection = ListSortDirection.Descending;
            FileNameSortDirection = ListSortDirection.Descending;
            CreateTimeSortDirection = ListSortDirection.Descending;
            UpdateTimeSortDirection = ListSortDirection.Descending;
            ClassNoSortDirection = ListSortDirection.Descending;
            SecrecyLevelSortDirection = ListSortDirection.Descending;
            FilePathSortDirection = ListSortDirection.Descending;
        }

        /// <summary>
        /// 表示中のアイテムをブロック毎に取得
        /// </summary>
        /// <returns>通常ファイルのリスト,Zip1,Zip2....</returns>
        private Dictionary<string, List<DataGridViewRow>>  GetDataGridViewRowDic(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = new Dictionary<string, List<DataGridViewRow>>();

            for (int i = 0; i <= (CurrentViewZipCount); i++)
            {
                string keyValue = i.ToString(); // 検索条件
                try
                {
                    var rows = from row in dataGridViewList.Rows.Cast<DataGridViewRow>()
                               where row.Cells["ColumnZipCount"].Value.ToString() == keyValue
                               select row;

                    List<DataGridViewRow> rowsList = new List<DataGridViewRow>(rows.ToList());

                    if (rowCollectionDic.ContainsKey(i.ToString()) == false)
                    {
                        rowCollectionDic[i.ToString()] = new List<DataGridViewRow>();
                    }

                    rowCollectionDic[i.ToString()] = rowsList;
                }
                catch (Exception ex)
                {
                    // 該当データなし時は、例外が発生するが処理として問題がないためスルーする
                }
            }

            return rowCollectionDic;
        }

        /// <summary>
        /// Zipフォーマットソート
        /// </summary>
        private void SortedZipFormat(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = GetDataGridViewRowDic(e);

            //並び替えの方向（昇順か降順か）を決める
            ZipSortDirection =
                    (ZipSortDirection == ListSortDirection.Ascending) ?
                    ListSortDirection.Descending : ListSortDirection.Ascending;

            dataGridViewList.Rows.Clear();

            if (ZipSortDirection == ListSortDirection.Descending)
            {
                for (int i = 0; i <= (CurrentViewZipCount); i++)
                {
                    if (rowCollectionDic.ContainsKey(i.ToString()) != false)
                    {
                        dataGridViewList.Rows.AddRange(rowCollectionDic[i.ToString()].ToArray());
                    }
                }
            }
            else
            {
                for (int i = (CurrentViewZipCount); i >= 0; i--)
                {
                    if (rowCollectionDic.ContainsKey(i.ToString()) != false)
                    {
                        dataGridViewList.Rows.AddRange(rowCollectionDic[i.ToString()].ToArray());
                    }
                }
            }
        }

        /// <summary>
        /// キーに対応するRowを取得する
        /// </summary>
        /// <param name="listRow">表示中のデータグリッド行</param>
        /// <param name="Key">キー</param>
        /// <returns>対応するRow</returns>
        private DataGridViewRow GetIndexRow(List<DataGridViewRow> listRow, string Key)
        {
            DataGridViewRow row = listRow.Where(x => string.IsNullOrEmpty(x.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString()) == false).Where(x =>
                        Path.Combine(x.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString(), x.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString())
                            == Key)
                            .FirstOrDefault();

            if (row == null)
            {
                row = listRow.Where(x =>
                        Path.Combine(x.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString(), x.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString())
                            == Key)
                            .FirstOrDefault();
            }

            return row;
        }

        /// <summary>
        /// インデックスからソート順を指定してグリッドにセットする
        /// </summary>
        /// <param name="direction">ソート方向</param>
        /// <param name="rowCollectionDic">データグリッド行リスト</param>
        /// <param name="IndexList">ソート用値</param>
        private void SetSortDataGrid(
                ListSortDirection direction,
                Dictionary<string, List<DataGridViewRow>> rowCollectionDic,
                Dictionary<string, string> IndexList
                )
        {
            List<KeyValuePair<string, string>> SortDicList = null;
            if (direction == ListSortDirection.Descending)
            {
                SortDicList = IndexList.OrderBy(x => x.Value).ToList();
            }
            else
            {
                SortDicList = IndexList.OrderByDescending(x => x.Value).ToList();
            }

            foreach (var SortDic in SortDicList)
            {
                foreach (var rowCollecton in rowCollectionDic)
                {
                    List<DataGridViewRow> listRow = rowCollecton.Value;
                    DataGridViewRow row = GetIndexRow(listRow, SortDic.Key);

                    if (row != null)
                    {
                        if (row.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString() == Resources.msgZipFormatFile)
                        {
                            // ZIPの場合一連のアイテムをセット
                            string ZipCount = row.Cells[(int)DataGridViewColumnIndex.ZipFileCount].Value.ToString();
                            dataGridViewList.Rows.AddRange(rowCollectionDic[ZipCount].ToArray());
                            break;

                        }
                        else
                        {
                            // 通常の場合普通にセット
                            dataGridViewList.Rows.Add(row);
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// ソート用キーをファイルパスから取得
        /// </summary>
        /// <param name="row">データグリッド行</param>
        /// <returns>ソート用キー</returns>
        private string GetIndexSortKey(DataGridViewRow row)
        {
            string FileName = "";
            if (string.IsNullOrEmpty(row.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString()) == false)
            {
                FileName = Path.Combine(row.Cells[(int)DataGridViewColumnIndex.ZipFilePath].Value.ToString(), row.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString());
            }
            else
            {
                FileName = Path.Combine(row.Cells[(int)DataGridViewColumnIndex.FilePath].Value.ToString(), row.Cells[(int)DataGridViewColumnIndex.FileName].Value.ToString());
            }

            return FileName;
        }

        /// <summary>
        /// 検索用ファイルパスをコレクションから取得する
        /// </summary>
        /// <param name="eIndex">列の検索列挙子</param>
        /// <param name="rowCollectionDic">データグリッド行</param>
        /// <returns>検索用ファイルパス</returns>
        private Dictionary<string, string> GetIndexList(
            DataGridViewColumnIndex eIndex,
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic
            )
        {
            Dictionary<string, string> IndexList = new Dictionary<string, string>();

            for (int i = 0; i <= (CurrentViewZipCount); i++)
            {
                if (rowCollectionDic.ContainsKey(i.ToString()) != false)
                {
                    List<DataGridViewRow> listRow = rowCollectionDic[i.ToString()];
                    if (i == 0)
                    {
                        foreach (DataGridViewRow row in listRow)
                        {
                            string FileName = GetIndexSortKey(row);
                            IndexList[FileName] =
                                row.Cells[(int)eIndex].Value.ToString();
                        }
                    }
                    else
                    {
                        string FileName = GetIndexSortKey(listRow[0]);
                        IndexList[FileName] =
                            listRow[0].Cells[(int)eIndex].Value.ToString();
                    }
                }
            }

            return IndexList;
        }

        /// <summary>
        /// ファイル名ソート
        /// </summary>
        private void SortedFileName(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = GetDataGridViewRowDic(e);

            //並び替えの方向（昇順か降順か）を決める
            FileNameSortDirection =
                    (FileNameSortDirection == ListSortDirection.Ascending) ?
                    ListSortDirection.Descending : ListSortDirection.Ascending;

            dataGridViewList.Rows.Clear();

            Dictionary<string, string> FileNameIndex = GetIndexList(
                DataGridViewColumnIndex.FileName,
                rowCollectionDic
                );

            // グリッドにセットする
            SetSortDataGrid(FileNameSortDirection, rowCollectionDic, FileNameIndex);
        }

        /// <summary>
        /// 作成日ソート
        /// </summary>
        private void SortedCreateTime(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = GetDataGridViewRowDic(e);

            //並び替えの方向（昇順か降順か）を決める
            CreateTimeSortDirection =
                    (CreateTimeSortDirection == ListSortDirection.Ascending) ?
                    ListSortDirection.Descending : ListSortDirection.Ascending;

            dataGridViewList.Rows.Clear();

            Dictionary<string, string> CreateTimeIndex = GetIndexList(
                    DataGridViewColumnIndex.CreatedDate,
                    rowCollectionDic
                    );

            // グリッドにセットする
            SetSortDataGrid(CreateTimeSortDirection, rowCollectionDic, CreateTimeIndex);
        }

        /// <summary>
        /// 更新日ソート
        /// </summary>
        private void SortedUpdateTime(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = GetDataGridViewRowDic(e);

            //並び替えの方向（昇順か降順か）を決める
            UpdateTimeSortDirection =
                    (UpdateTimeSortDirection == ListSortDirection.Ascending) ?
                    ListSortDirection.Descending : ListSortDirection.Ascending;

            dataGridViewList.Rows.Clear();

            Dictionary<string, string> UpdateTimeIndex = GetIndexList(
                    DataGridViewColumnIndex.UpdatedDate,
                    rowCollectionDic
                    );

            SetSortDataGrid(UpdateTimeSortDirection, rowCollectionDic, UpdateTimeIndex);
        }

        /// <summary>
        /// 文書分類ソートパラメータ
        /// </summary>
        /// <param name="x">比較対象行</param>
        /// <param name="y">比較対象行</param>
        /// <returns>優先順位</returns>
        private static int CompareDocumentType(DataGridViewRow x, DataGridViewRow y)
        {
            string keyA = x.Cells[(int)DataGridViewColumnIndex.DocumentType].Value.ToString();
            string keyB = y.Cells[(int)DataGridViewColumnIndex.DocumentType].Value.ToString();

            int returnLevel = 0;

            if (ClassNoSortDirection == ListSortDirection.Descending)
            {
                returnLevel = -(string.Compare(keyA, keyB));
            }
            else
            {
                returnLevel = (string.Compare(keyA, keyB));
            }

            return returnLevel;
        }

        /// <summary>
        /// 文書分類ソート
        /// </summary>
        private void SortedClassNo(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = GetDataGridViewRowDic(e);

            //並び替えの方向（昇順か降順か）を決める
            ClassNoSortDirection =
                    (ClassNoSortDirection == ListSortDirection.Ascending) ?
                    ListSortDirection.Descending : ListSortDirection.Ascending;

            dataGridViewList.Rows.Clear();

            for (int i = 0; i <= (CurrentViewZipCount); i++)
            {
                if (rowCollectionDic.ContainsKey(i.ToString()) != false)
                {
                    List<DataGridViewRow> listRow = rowCollectionDic[i.ToString()];

                    if (i == 0)
                    {
                        // 通常ファイル
                        if (ClassNoSortDirection == ListSortDirection.Descending)
                        {
                            var ToList = listRow.OrderBy(row => row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value.ToString())
                                .ToList();

                            ToList.Sort(CompareDocumentType);
                            dataGridViewList.Rows.AddRange(ToList.ToArray());
                        }
                        else
                        {
                            var ToList = listRow.OrderByDescending(row => row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value.ToString())
                                  .ToList();

                            ToList.Sort(CompareDocumentType);
                            dataGridViewList.Rows.AddRange(ToList.ToArray());
                        }
                    }
                    else
                    {
                        // ZIPファイル系
                        if (ClassNoSortDirection == ListSortDirection.Descending)
                        {

                            var Header = listRow[0];
                            var ToArray = listRow.
                                Where(row => row.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString() != Resources.msgZipFormatFile)
                                .OrderBy(row => row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value.ToString())
                                .ToList();

                            ToArray.Sort(CompareDocumentType);

                            dataGridViewList.Rows.Add(Header);
                            dataGridViewList.Rows.AddRange(ToArray.ToArray());
                        }
                        else
                        {
                            var Header = listRow[0];
                            var ToArray = listRow.
                                Where(row => row.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString() != Resources.msgZipFormatFile)
                                .OrderByDescending(row => row.Cells[(int)DataGridViewColumnIndex.DocumentType].Value.ToString())
                                .ToList();

                            ToArray.Sort(CompareDocumentType);

                            dataGridViewList.Rows.Add(Header);
                            dataGridViewList.Rows.AddRange(ToArray.ToArray());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 機密区分ソートパラメータ
        /// </summary>
        /// <param name="x">比較対象行</param>
        /// <param name="y">比較対象行</param>
        /// <returns>優先順位</returns>
        private static int CompareSecretType(DataGridViewRow x, DataGridViewRow y)
        {
            string keyA = x.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString();
            string keyB = y.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString();

            int KeyACount = 0;
            int KeyBCount = 0;

            int SecrecyS = 4;
            int SecrecyA = 3;
            int SecrecyB = 2;
            int SecrecyDefault = 1;
            int SecrecyError = 0;

            if (SecrecyLevelSortDirection != ListSortDirection.Descending)
            {
                SecrecyS = 0;
                SecrecyA = 1;
                SecrecyB = 2;
                SecrecyDefault = 3;
                SecrecyError = 4;
            }

            switch (keyA)
            {
                case SECRECY_PROPERTY_S:
                    KeyACount = SecrecyS;
                    break;
                case SECRECY_PROPERTY_A:
                    KeyACount = SecrecyA;
                    break;
                case SECRECY_PROPERTY_B:
                    KeyACount = SecrecyB;
                    break;
                case "":
                    KeyACount = SecrecyError;
                    break;
                default:
                    KeyACount = SecrecyDefault;
                    break;
            }

            switch (keyB)
            {
                case SECRECY_PROPERTY_S:
                    KeyBCount = SecrecyS;
                    break;
                case SECRECY_PROPERTY_A:
                    KeyBCount = SecrecyA;
                    break;
                case SECRECY_PROPERTY_B:
                    KeyBCount = SecrecyB;
                    break;
                case "":
                    KeyBCount = SecrecyError;
                    break;
                default:
                    KeyBCount = SecrecyDefault;
                    break;
            }

            if (KeyACount > KeyBCount)
            {
                return 1;
            }
            else if (KeyACount < KeyBCount)
            {
                return -1;
            }

            return 0;
        }

        /// <summary>
        /// 機密区分ソート
        /// </summary>
        private void SortedSecrecyLevel(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = GetDataGridViewRowDic(e);

            //並び替えの方向（昇順か降順か）を決める
            SecrecyLevelSortDirection =
                    (SecrecyLevelSortDirection == ListSortDirection.Ascending) ?
                    ListSortDirection.Descending : ListSortDirection.Ascending;

            dataGridViewList.Rows.Clear();

            for (int i = 0; i <= (CurrentViewZipCount); i++)
            {
                if (rowCollectionDic.ContainsKey(i.ToString()) != false)
                {
                    List<DataGridViewRow> listRow = rowCollectionDic[i.ToString()];

                    if (i == 0)
                    {
                        // 通常ファイル
                        if (SecrecyLevelSortDirection == ListSortDirection.Descending)
                        {
                            var ToList = listRow.OrderBy(row => row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString())
                                .ToList();

                            ToList.Sort(CompareSecretType);
                            dataGridViewList.Rows.AddRange(ToList.ToArray());
                        }
                        else
                        {
                            var ToList = listRow.OrderByDescending(row => row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString())
                                  .ToList();

                            ToList.Sort(CompareSecretType);
                            dataGridViewList.Rows.AddRange(ToList.ToArray());
                        }
                    }
                    else
                    {
                        // ZIPファイル系
                        if (SecrecyLevelSortDirection == ListSortDirection.Descending)
                        {

                            var Header = listRow[0];
                            var ToArray = listRow.
                                Where(row => row.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString() != Resources.msgZipFormatFile)
                                .OrderBy(row => row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString())
                                .ToList();


                            ToArray.Sort(CompareSecretType);

                            dataGridViewList.Rows.Add(Header);
                            dataGridViewList.Rows.AddRange(ToArray.ToArray());
                        }
                        else
                        {
                            var Header = listRow[0];
                            var ToArray = listRow.
                                Where(row => row.Cells[(int)DataGridViewColumnIndex.ZipFormat].Value.ToString() != Resources.msgZipFormatFile)
                                .OrderByDescending(row => row.Cells[(int)DataGridViewColumnIndex.SecretType].Value.ToString())
                                .ToList();


                            ToArray.Sort(CompareSecretType);

                            dataGridViewList.Rows.Add(Header);
                            dataGridViewList.Rows.AddRange(ToArray.ToArray());
                        }
                    }
                }
            }
        }

        /// <summary>
        /// ファイルパスソート
        /// </summary>
        private void SortedFilePath(DataGridViewCellMouseEventArgs e)
        {
            Dictionary<string, List<DataGridViewRow>> rowCollectionDic = GetDataGridViewRowDic(e);

            //並び替えの方向（昇順か降順か）を決める
            FilePathSortDirection =
                    (FilePathSortDirection == ListSortDirection.Ascending) ?
                    ListSortDirection.Descending : ListSortDirection.Ascending;

            dataGridViewList.Rows.Clear();

            Dictionary<string, string> FilePathIndex = GetIndexList(
                    DataGridViewColumnIndex.FilePath,
                    rowCollectionDic
                    );

            SetSortDataGrid(FilePathSortDirection, rowCollectionDic, FilePathIndex);
        }

        #endregion

    }
}
