using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using DocumentManagement_List.Properties;   // step2 iwasa

namespace DocumentManagement_List
{
    public partial class FormSetting : Form
    {
        #region <定数定義>
        /// <summary>
        /// 定数定義
        /// </summary>

        // 画面関連
        public const string PROCESS_VIEW_FORMAT = "{0} / {1} 件　処理中";

        // 履歴関連
        public const int MAX_SEARCH_HISTORY_COUNT = 10;             // 最大検索履歴数
        public const int MAX_HISTORY_COUNT = 10;                    // 最大履歴数

        // ファイルプロパティ関連
        public const string SECRECY_S = "S秘";                      // SAB秘 S秘
        public const string SECRECY_A = "A秘";                      // SAB秘 A秘
        public const string SECRECY_B = "B秘";                      // SAB秘 B秘
        public const string SECRECY_PROPERTY_S = "SecrecyS";        // プロパティに書き込むSAB秘 S秘
        public const string SECRECY_PROPERTY_A = "SecrecyA";        // プロパティに書き込むSAB秘 A秘
        public const string SECRECY_PROPERTY_B = "SecrecyB";        // プロパティに書き込むSAB秘 B秘
        public const string NO_STAMP = "no_stamp";                  // プロパティに書き込むスタンプ無し情報

        // 201701009 追加
        // 他事業所モード
        public const string OTHER_OFFICE_MODE2 = "2";                 // 他事業所モード２（他事業所のドキュメントの場合は表示画面を表示する）
        public const string OTHER_OFFICE_MODE3 = "3";                 // 他事業所モード３（他事業所のドキュメントの場合は表示画面を表示しない）

        // 共通設定関連
        public const string COMMON_SETFOLDERNAME = "SAB";           // 20171009 追加（共通設定格納フォルダ名）
        public const string COMMON_SETFILENAME = "common_setting.config";
                                                                    // 共通設定ファイル名
        public const string COMMON_SETDEF_CLASSSSETTING_PATH = @"C:\Program Files\SAB\pic\属性設定.csv";
                                                                    // 共通設定ファイルデフォルト属性設定CSVファイルパス
        public const string COMMON_SETDEF_IMAGE_FOLDER_PATH = @"C:\Program Files\SAB\pic";
                                                                    // 共通設定ファイルデフォルトイメージ画像フォルダパス
        public const string COMMON_SETDEF_HMDOCUMENT_PATH = @"C:\Program Files\SAB\other\HM文書管理規程(付表).xlsx";
                                                                    // 共通設定ファイルデフォルトHM文書管理規程ファイルパス
        public const string COMMON_SETDEF_SKIPMODE = "";            // 共通設定ファイルデフォルトスキップモード
        public const string COMMON_SETDEF_CLASSNO = "--.99.9999";   // 共通設定ファイルデフォルトスキップモード文書分類番号
        public const string COMMON_SETDEF_SECLV = SECRECY_PROPERTY_B;
                                                                    // 共通設定ファイルデフォルトスキップモード機密区分

        // 20200804 追加
        public const string COMMON_SETDEF_TEMP_PATH = @"C:\tmp\unzip"; // 一時保管場所

        // 20171009 追加（他事業所対応）
        public const string COMMON_SETDEF_OFFICECODE = "HM";        // 共通設定事業所コード
        public const string COMMON_SETDEF_OTHEROFFICEMODE = OTHER_OFFICE_MODE2;    // 共通設定他事業所モード

        // ユーザー設定関連
        //public const string USER_SETFILENAME = "user_setting.config";  // ユーザー設定ファイル名
        public const string USER_SETFILENAME = "user_setting.config";  // ユーザー設定ファイル名
        public const string USER_SETFOLDERNAME = "SAB";             // 20171009 追加 ユーザー設定格納フォルダ名（新）
        public const string USER_SET_OLDFOLDERNAME = @"Microsoft Corporation\Microsoft Office 2010";  // 20171009 追加 ユーザー設定格納フォルダ名（旧）
        #endregion

        #region <内部変数>
        /// <summary>
        /// 内部変数
        /// </summary>
        private ClassAttributeSetting clsModule;                    // 属性設定情報クラス
        private UserSettings clsUserSettting;                       // ユーザー設定クラス
        public CommonSettings clsCommonSettting;                    // 共通設定クラス
        public string strFilePropertyClassNo;                       // ファイルプロパティ情報 文書分類番号
        public string strFilePropertySecrecyLevel;                  // ファイルプロパティ情報 機密区分
        public string strFilePropertyOfficeCode;                    // 20171009 ファイルプロパティ情報 事業所コード
        public Boolean bFilePropertyStamp;                          // ファイルプロパティ情報 スタンプ有無
        public Boolean bCommonError;                                // 共通設定エラーフラグ
        public string StrClassNo;                                   // 文書分類番号
        public string StrSecrecyLevel;                              // 機密区分
        private Boolean BFormOkCancel;                              // 登録、後で登録ボタン押下フラグ
        public string StrTargetFilePath;                            // 対象ファイルパス(フルパス)
        public int IProcessCnt;                                     // 処理件数
        public int ITotalProcessCnt;                                // 全体処理件数
        #endregion

        #region <クラス定義>
        /// <summary>
        /// クラス定義
        /// </summary>

        // ユーザー設定
        public class UserSettings
        {
            public string strDefaultLargeClassName;                 // デフォルト大分類初期値
            public string strDefaultMiddleClassName;                // デフォルト中分類初期値
            public string strDefaultSmallClassName;                 // デフォルト小分類初期値
            public string[] strSearchHistory;                       // 検索履歴
            public string[] strHistoryLargeClassName;               // 履歴大分類
            public string[] strHistoryMiddleClassName;              // 履歴中分類
            public string[] strHistorySmallClassName;               // 履歴小分類
            public string[] strHistoryDocumentExample;              // 履歴文書例

            public UserSettings()
            {
                strDefaultLargeClassName = "";                      // デフォルト大分類初期化
                strDefaultMiddleClassName = "";                     // デフォルト中分類初期化
                strDefaultSmallClassName = "";                      // デフォルト小分類初期化

                // 検索履歴初期化
                strSearchHistory = new string[MAX_SEARCH_HISTORY_COUNT];
                for (int iCnt = 0; iCnt < MAX_SEARCH_HISTORY_COUNT; iCnt++)
                {
                    strSearchHistory[iCnt] = "";
                }

                // 履歴初期化
                strHistoryLargeClassName = new string[MAX_HISTORY_COUNT];
                strHistoryMiddleClassName = new string[MAX_HISTORY_COUNT];
                strHistorySmallClassName = new string[MAX_HISTORY_COUNT];
                strHistoryDocumentExample = new string[MAX_HISTORY_COUNT];
                for (int iCnt = 0; iCnt < MAX_HISTORY_COUNT; iCnt++)
                {
                    strHistoryLargeClassName[iCnt] = "";
                    strHistoryMiddleClassName[iCnt] = "";
                    strHistorySmallClassName[iCnt] = "";
                    strHistoryDocumentExample[iCnt] = "";
                }
            }
        }

        // 共通設定
        public class CommonSettings
        {
            public string strClassSettingFilePath;                  // 属性設定CSVファイルパス
            public string strImageFolderPath;                       // イメージ画像フォルダパス
            public string strHMDocumentFilePath;                    // HM文書管理規程ファイルパス
            public string strExcelSkipMode;                         // Excelスキップモード(3の場合は保存動作時はデフォルト設定で登録)(Excelのみ有効)
            public string strDefaultClassNo;                        // Excelスキップモードデフォルト文書分類番号
            public string strDefaultSecrecyLevel;                   // Excelスキップモードデフォルト機密区分

            // 20171009 追加（他事業所対応）
            public string strOfficeCode;                            // 事業所コード
            public string strOtherOfficeMode;                       // 他事業所処理モード

            // 20200804 追加 ()
            public string strTempPath;                              // 一時保管場所
            public List<string> lstSecureFolder;                    // セキュアフォルダリスト

            // step2
            // 文書のローカルパス
            public string strSABListLocalPath;

            // 文書のサーバーパス
            public string strSABListServerPath;

            // step2 iwasa
            // 言語設定
            public string strCulture;

            // 最終版の文言
            public List<string> lstFinal;

            public CommonSettings()
            {
                strClassSettingFilePath = COMMON_SETDEF_CLASSSSETTING_PATH;
                                                                    // 属性設定CSVファイルパス初期化
                strImageFolderPath = COMMON_SETDEF_IMAGE_FOLDER_PATH;
                                                                    // イメージ画像フォルダパス初期化
                strHMDocumentFilePath = COMMON_SETDEF_HMDOCUMENT_PATH;
                                                                    // HM文書管理規程ファイルパス初期化
                strExcelSkipMode = COMMON_SETDEF_SKIPMODE;          // Excelスキップモード初期化
                strDefaultClassNo = COMMON_SETDEF_CLASSNO;          // Excelスキップモードデフォルト文書分類番号初期化
                strDefaultSecrecyLevel = COMMON_SETDEF_SECLV;       // Excelスキップモードデフォルト機密区分初期化

                // 20171009 追加（他事業所対応）
                strOfficeCode = COMMON_SETDEF_OFFICECODE;           // 事業所コード初期化
                strOtherOfficeMode = COMMON_SETDEF_OTHEROFFICEMODE; // 他事業所処理モード初期化

                // 20200804 追加
                strTempPath = COMMON_SETDEF_TEMP_PATH;              // 一時保管場所パス初期化
                lstSecureFolder = new List<string>();               // セキュアフォルダリスト初期化
                //lstSecureFolder.Add(@"C:\tmp\SecureA");
                //lstSecureFolder.Add(@"C:\tmp\SecureS");
                //lstSecureFolder.Add(@"C:\tmp\SecureSS");

                // step2
                // 文書のローカルパス
                strSABListLocalPath = "";
                //strSABListLocalPath = @"C:\HLI\local\document.xlsx";

                // 文書のサーバーパス
                strSABListServerPath = "";
                //strSABListServerPath = @"C:\HLI\network\fileServer\document.xlsx";

                // 言語コード初期化 step2 iwasa
                strCulture = System.Threading.Thread.CurrentThread.CurrentUICulture.ToString();

                // 最終版の文言
                lstFinal = new List<string>();
            }
        }

        #endregion
        
        #region <コンストラクタ>
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FormSetting()
        {
            // 共通設定読み込み
            clsCommonSettting = new CommonSettings();

            // 共通設定エラー時処理
            bCommonError = CommonSettingRead();
            if (bCommonError == false)
            {
                //MessageBox.Show("共通設定ファイルの読み込みに失敗しました。");
                MessageBox.Show(Resources.msgFailedReadCommonFile); // step2 iwasa
            }
            else
            {
                // ユーザー設定読み込み
                clsUserSettting = new UserSettings();
                if (UserSettingRead() == false)
                {
                    //MessageBox.Show("ユーザー設定ファイルの読み込みに失敗しました。");
                    MessageBox.Show(Resources.msgFailedReadUserSetting);    // step2 iwasa
                }

                // CSVファイル読み込み step2 属性設定.csvが不要のためコメントアウト
                //clsModule = new ClassAttributeSetting();
                //if (clsModule.readCsvFile(clsCommonSettting.strClassSettingFilePath) == false) // 2016/09/20エラーチェック処理追加
                //{
                //    //MessageBox.Show("CSVファイルの読み込みに失敗しました。");
                //    //MessageBox.Show(Resources.msgFailedReadCsvFile);    // step2 iwasa
                //}
            }

            // 各コンポーネント初期化
            InitializeComponent();
        }
        #endregion

        #region <フォームイベント>
        /// <summary>
        /// フォームロード
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormSetting_Load(object sender, EventArgs e)
        {
            // エラー時の場合は画面を閉じる
            if (bCommonError == false)
            {
                this.Close();
                return;
            }

            BFormOkCancel = false;

            #region<各分類初期化>
            // 表示項目クリア 2016/09/20追加
            comboBoxLargeClass.Items.Clear();
            comboBoxMiddleClass.Items.Clear();
            comboBoxSmallClass.Items.Clear();
            listBoxDocumentExample.Items.Clear();
            labelClassNo.Text = "";
            labelSmallClass.Text = "";
            labelSecrecyLevel.Text = "";

            // 既存の文書分類番号を分解
            string strBuf = (string.IsNullOrEmpty(strFilePropertyClassNo)) ? "" : strFilePropertyClassNo;
            string[] stProperty = strBuf.Split('.');

            // 大分類名コンボボックス初期設定
            string[] strLargeClassList = clsModule.getLargeClassNameList();

            //// エラーログファイル名「errorLog_yyyymmddHHMMss.log」
            //using (StreamWriter writer = new StreamWriter(ListForm.EXCLUSION_LOG_PATH + @"\Log" + @"\debug.log", true))    // 末尾に追記
            //{
            //    writer.WriteLine(DateTime.Now.ToString() + " □□□□□□□□□□ strLargeClassList □□□□□□□□□ ");
            //    // エラーあり、ファイル出力
            //    foreach (string err in strLargeClassList)
            //    {
            //        // エラーログファイルに書込（末尾に追記）
            //        writer.WriteLine(err);
            //    }
            //    writer.Close();
            //}

            comboBoxLargeClass.Items.Clear();
            foreach(string strLargeClass in strLargeClassList)
            {
                comboBoxLargeClass.Items.Add(strLargeClass);
            }
            // 大分類初期値設定
            if(string.IsNullOrEmpty(strFilePropertyClassNo) == false)
            {
                // 初期設定
                comboBoxLargeClass.Text = clsUserSettting.strDefaultLargeClassName;

                // 文書分類番号が設定済みの場合は設定されている値を選択する
                if (stProperty.Count() == 3)
                {
                    foreach (string strClass in comboBoxLargeClass.Items)
                    {
                        if (strClass.StartsWith(stProperty[0]))
                        {
                            comboBoxLargeClass.Text = strClass;
                            break;
                        }
                    }
                }
            }
            else if (string.IsNullOrEmpty(clsUserSettting.strDefaultLargeClassName) == true)
            {
                if (comboBoxLargeClass.Items.Count > 0) comboBoxLargeClass.SelectedIndex = 0; // 2016/09/20大分類項目がなかった場合のエラーで落ちる問題を修正
            }
            else
            {
                comboBoxLargeClass.Text = clsUserSettting.strDefaultLargeClassName;
            }

            // 中分類名コンボボックス初期設定
            string[] strMiddleClassList = clsModule.getMiddleClassNameList(comboBoxLargeClass.Text);
            comboBoxMiddleClass.Items.Clear();
            foreach (string strMiddleClass in strMiddleClassList)
            {
                comboBoxMiddleClass.Items.Add(strMiddleClass);
            }
            // 中分類初期値設定
            if (string.IsNullOrEmpty(strFilePropertyClassNo) == false)
            {
                // 初期設定
                comboBoxMiddleClass.Text = clsUserSettting.strDefaultMiddleClassName;

                // 文書分類番号が設定済みの場合は設定されている値を選択する
                if (stProperty.Count() == 3)
                {
                    foreach (string strClass in comboBoxMiddleClass.Items)
                    {
                        if (strClass.StartsWith(stProperty[1]))
                        {
                            comboBoxMiddleClass.Text = strClass;
                            break;
                        }
                    }
                }
            }
            else if (string.IsNullOrEmpty(clsUserSettting.strDefaultMiddleClassName) == true)
            {
                if (comboBoxMiddleClass.Items.Count > 0) comboBoxMiddleClass.SelectedIndex = 0;
            }
            else
            {
                comboBoxMiddleClass.Text = clsUserSettting.strDefaultMiddleClassName;
            }

            // 小分類名コンボボックス初期設定
            string[] strSmallClassList = clsModule.getSmallClassNameList(comboBoxLargeClass.Text, comboBoxMiddleClass.Text);
            comboBoxSmallClass.Items.Clear();
            foreach (string strSmallClass in strSmallClassList)
            {
                comboBoxSmallClass.Items.Add(strSmallClass);
            }
            // 小分類初期値設定
            if (string.IsNullOrEmpty(strFilePropertyClassNo) == false)
            {
                // 初期設定
                comboBoxSmallClass.Text = clsUserSettting.strDefaultSmallClassName;

                // 文書分類番号が設定済みの場合は設定されている値を選択する
                if (stProperty.Count() == 3)
                {
                    foreach (string strClass in comboBoxSmallClass.Items)
                    {
                        if (strClass.StartsWith(stProperty[2]))
                        {
                            comboBoxSmallClass.Text = strClass;
                            break;
                        }
                    }
                }
            }
            else if (string.IsNullOrEmpty(clsUserSettting.strDefaultSmallClassName) == true)
            {
                if (comboBoxSmallClass.Items.Count > 0) comboBoxSmallClass.SelectedIndex = 0;
            }
            else
            {
                comboBoxSmallClass.Text = clsUserSettting.strDefaultSmallClassName;
            }
            #endregion

            if (string.IsNullOrEmpty(StrTargetFilePath))
            {
                // 処理対象パスがない場合(一括設定)

                // 処理件数表示
                labelProgress.Text = "";

                // 処理対象ファイルパス表示
                labelFilePath.Text = "";
            }
            else
            {
                // 処理対象パスがある場合(個別設定)

                // 処理件数表示
                labelProgress.Text = string.Format(PROCESS_VIEW_FORMAT, IProcessCnt, ITotalProcessCnt);

                // 処理対象ファイルパス表示
                labelFilePath.Text = StrTargetFilePath;
            }

            // 登録ボタンにフォーカスをセット
            buttonRegist.Focus();
        }

        /// <summary>
        /// フォームクローズ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormSetting_FormClosing(object sender, FormClosingEventArgs e)
        {
            // フォーム設定
            if (!BFormOkCancel)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.No;
            }
        }

        /// <summary>
        /// 後で登録するボタンクリック時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonNotRegist_Click(object sender, EventArgs e)
        {
            // ダイアログ結果キャンセル設定
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            BFormOkCancel = true;

            // フォームを閉じる
            this.Close();
        }

        /// <summary>
        /// フォームキーダウン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormSetting_KeyDown(object sender, KeyEventArgs e)
        {
            // ESCキーが押された場合
            if (e.KeyData == Keys.Escape)
            {
                // 後で登録ボタンクリック
                buttonNotRegist_Click(sender, e);
            }
            // Enterキーが押された場合
            else if (e.KeyData == Keys.Enter)
            {
                // 登録ボタンクリック処理
                buttonRegist_Click(sender, e);
            }
        }

        /// <summary>
        /// 登録履歴ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonHistory_Click(object sender, EventArgs e)
        {
            // 履歴画面設定
            HistoryForm frmHistory = new HistoryForm();
            frmHistory.strClassSettingFilePath = clsCommonSettting.strClassSettingFilePath;
            frmHistory.StartPosition = FormStartPosition.CenterScreen;

            // 履歴を渡す
            frmHistory.strHistoryLargeClassName = clsUserSettting.strHistoryLargeClassName;
            frmHistory.strHistoryMiddleClassName = clsUserSettting.strHistoryMiddleClassName;
            frmHistory.strHistorySmallClassName = clsUserSettting.strHistorySmallClassName;
            frmHistory.strHistoryDocumentExample = clsUserSettting.strHistoryDocumentExample;

            // 画面表示
            if (frmHistory.ShowDialog() == DialogResult.OK)
            {
                // 項目が選択された場合はコンボボックスを更新
                ComboBoxClassSetting(frmHistory.strSelectLargeClassName, frmHistory.strSelectMiddleClassName, frmHistory.strSelectSmallClassName);

                // 文書例選択
                for(int iCnt=0 ; iCnt < listBoxDocumentExample.Items.Count ; iCnt++)
                {
                    if (listBoxDocumentExample.Items[iCnt].ToString() == frmHistory.strSelectDocumentExample)
                    {
                        listBoxDocumentExample.SelectedIndex = iCnt;
                    }
                }
            }

        }

        /// <summary>
        /// 文書分類検索ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSearch_Click(object sender, EventArgs e)
        {
            // 検索画面設定
            SearchForm frmSearch = new SearchForm();
            frmSearch.strClassSettingFilePath = clsCommonSettting.strClassSettingFilePath;
            frmSearch.StartPosition = FormStartPosition.CenterScreen;

            // 検索履歴を渡す
            frmSearch.strSearchHistory = clsUserSettting.strSearchHistory;
            
            // 画面表示
            if (frmSearch.ShowDialog() == DialogResult.OK)
            {
                int iCnt;

                // 項目が選択された場合はコンボボックスを更新
                ComboBoxClassSetting(frmSearch.strSelectLargeClassName, frmSearch.strSelectMiddleClassName, frmSearch.strSelectSmallClassName);

                // 文書例選択
                for (iCnt = 0; iCnt < listBoxDocumentExample.Items.Count; iCnt++)
                {
                    if (listBoxDocumentExample.Items[iCnt].ToString() == frmSearch.strSelectDocumentExample)
                    {
                        listBoxDocumentExample.SelectedIndex = iCnt;
                    }
                }

                // 検索履歴に同じワードがないか検索する
                for (iCnt = 0; iCnt < MAX_SEARCH_HISTORY_COUNT; iCnt++)
                {
                    if (clsUserSettting.strSearchHistory[iCnt] == frmSearch.strSearchWord) break;
                }

                // １つずつ下にずらす
                if (iCnt >= MAX_SEARCH_HISTORY_COUNT) iCnt = MAX_SEARCH_HISTORY_COUNT - 1;
                for (; iCnt > 0; iCnt--)
                {
                    clsUserSettting.strSearchHistory[iCnt] = clsUserSettting.strSearchHistory[iCnt - 1];
                }

                // 先頭に今回検索したワードを登録する
                clsUserSettting.strSearchHistory[0] = frmSearch.strSearchWord;
                
                // 設定ファイル書き込み
                if (UserSettingWrite() == false)
                {
                    MessageBox.Show("ユーザー設定ファイルの書き込みに失敗しました。");
                }
            }
        }

        /// <summary>
        /// 大分類コンボボックス変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxLargeClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 項目が選択された場合はコンボボックスを更新
            ComboBoxClassSetting(comboBoxLargeClass.Text, null, null);
        }

        /// <summary>
        /// 中分類コンボボックス変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxMiddleClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 項目が選択された場合はコンボボックスを更新
            ComboBoxClassSetting(null, comboBoxMiddleClass.Text, null);
        }

        /// <summary>
        /// 小分類コンボボックス変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxSmallClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 項目が選択された場合はコンボボックスを更新
            ComboBoxClassSetting(null, null, comboBoxSmallClass.Text);
        }

        /// <summary>
        /// 文書の名称（サンプル）選択処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxDocumentExample_SelectedIndexChanged(object sender, EventArgs e)
        {
            // SAB秘設定変更時処理
            SABSetting();
        }

        /// <summary>
        /// 初期値設定ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonDefaultSetting_Click(object sender, EventArgs e)
        {
            // 画面の内容をユーザー設定に反映
            clsUserSettting.strDefaultLargeClassName = comboBoxLargeClass.Text;
            clsUserSettting.strDefaultMiddleClassName = comboBoxMiddleClass.Text;
            clsUserSettting.strDefaultSmallClassName = comboBoxSmallClass.Text;

            // ユーザー設定書き込み
            if(UserSettingWrite() == false)
            {
                MessageBox.Show("ユーザー設定の書き込みに失敗しました。");
            }
        }

        /// <summary>
        /// 登録ボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRegist_Click(object sender, EventArgs e)
        {
            string strClassNo = labelClassNo.Text;                  // 文書分類番号
            string strSecrecyLevel;                                 // 機密区分
            string strFilePropertyOfficeCode;                       // 20171009追加（ファイル事業所コード）

            #region<入力値チェック>
            // 大項目～中項目選択チェック
            if ((comboBoxLargeClass.SelectedIndex < 0) || (comboBoxMiddleClass.SelectedIndex < 0)) return;
            #endregion

            // 機密区分設定
            switch (labelSecrecyLevel.Text)
            {
                case SECRECY_S:
                    strSecrecyLevel = SECRECY_PROPERTY_S;
                    break;
                case SECRECY_A:
                    strSecrecyLevel = SECRECY_PROPERTY_A;
                    break;
                case SECRECY_B:
                    strSecrecyLevel = SECRECY_PROPERTY_B;
                    break;
                default:
                    strSecrecyLevel = labelSecrecyLevel.Text;
                    break;
            }

            // 中分類が「仕掛中・作成中」などコードが99の場合はスタンプを押さない
            string[] strMiddleCodeStrit = comboBoxMiddleClass.Text.Trim().Split('.');
            string strMiddleCode = (strMiddleCodeStrit.Count() > 1) ? strMiddleCodeStrit[0] : "99";

            #region<履歴更新>
            // 登録履歴に同じワードがないか検索する
            int iCnt;
            for (iCnt = 0; iCnt < MAX_SEARCH_HISTORY_COUNT; iCnt++)
            {
                if ((clsUserSettting.strHistoryLargeClassName[iCnt] == comboBoxLargeClass.Text) &&
                    (clsUserSettting.strHistoryMiddleClassName[iCnt] == comboBoxMiddleClass.Text) &&
                    (clsUserSettting.strHistorySmallClassName[iCnt] == comboBoxSmallClass.Text) &&
                    (clsUserSettting.strHistoryDocumentExample[iCnt] == listBoxDocumentExample.SelectedItem.ToString())) break;
            }

            // １つずつ下にずらす
            if (iCnt >= MAX_SEARCH_HISTORY_COUNT) iCnt = MAX_SEARCH_HISTORY_COUNT - 1;
            for (; iCnt > 0; iCnt--)
            {
                clsUserSettting.strHistoryLargeClassName[iCnt] = clsUserSettting.strHistoryLargeClassName[iCnt - 1];
                clsUserSettting.strHistoryMiddleClassName[iCnt] = clsUserSettting.strHistoryMiddleClassName[iCnt - 1];
                clsUserSettting.strHistorySmallClassName[iCnt] = clsUserSettting.strHistorySmallClassName[iCnt - 1];
                clsUserSettting.strHistoryDocumentExample[iCnt] = clsUserSettting.strHistoryDocumentExample[iCnt - 1];
            }

            // 先頭に今回検索したワードを登録する
            clsUserSettting.strHistoryLargeClassName[0] = comboBoxLargeClass.Text;
            clsUserSettting.strHistoryMiddleClassName[0] = comboBoxMiddleClass.Text;
            clsUserSettting.strHistorySmallClassName[0] = comboBoxSmallClass.Text;
            clsUserSettting.strHistoryDocumentExample[0] = listBoxDocumentExample.SelectedItem.ToString();

            // 設定ファイル書き込み
            if (UserSettingWrite() == false)
            {
                MessageBox.Show("ユーザー設定ファイルの書き込みに失敗しました。");
            }
            #endregion

            // 20170801 修正（機密ランクダウン時に警告表示）
            // 変更前の機密ランクが「S」or「A」の場合
            if (strFilePropertyClassNo == SECRECY_PROPERTY_S || strFilePropertyClassNo == SECRECY_PROPERTY_A)
            {
                // 変更前の機密ランクが「S」の場合
                if (strFilePropertyClassNo == SECRECY_PROPERTY_S)
                {
                    // 変更後の機密ランクが「S」以外の場合
                    if(strClassNo != SECRECY_PROPERTY_S)
                    {
                        // 警告表示
                        MessageBox.Show("機密ランクが" + strFilePropertyClassNo + "から" + strClassNo + "へ下がりました。");
                    }
                }
                // 変更前の機密ランクが「A」の場合
                if (strFilePropertyClassNo == SECRECY_PROPERTY_A)
                {
                    // 変更後の機密ランクが「S」「A」以外の場合
                    if (!(strClassNo == SECRECY_PROPERTY_S || strClassNo == SECRECY_PROPERTY_A))
                    {
                        // 警告表示
                        MessageBox.Show("機密ランクが" + strFilePropertyClassNo + "から" + strClassNo + "へ下がりました。");
                    }
                }
            }

            // 変更前の機密ランク保持用


            // 属性設定反映
            StrClassNo = strClassNo;
            StrSecrecyLevel = strSecrecyLevel;

            // 20170801 コメントアウト（機密ランクダウン時に警告表示）
            //// 属性設定反映
            //StrClassNo = strClassNo;
            //StrSecrecyLevel = strSecrecyLevel;

            // ダイアログを閉じる
            BFormOkCancel = true;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        /// <summary>
        /// HM文書管理規程ラベルクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void labelHMDocumentManagement_Click(object sender, EventArgs e)
        {
            // 拡張子取得
            System.IO.FileInfo cFileInfo = new System.IO.FileInfo(clsCommonSettting.strHMDocumentFilePath);

            // HM文書管理規程ドキュメントを開く（現在開いているアプリと同じ形式のファイルを開こうとすると固まるためEXE指定し別ウィンドウで実行する）
            switch (cFileInfo.Extension)
            {
                case ".xls":
                case ".xlsx":
                    // Excel
                    System.Diagnostics.Process.Start("EXCEL.EXE", '"' + clsCommonSettting.strHMDocumentFilePath + '"');
                    break;

                case ".doc":
                case ".docx":
                    // Word
                    System.Diagnostics.Process.Start("WINWORD.EXE", '"' + clsCommonSettting.strHMDocumentFilePath + '"');
                    break;

                case ".ppt":
                case ".pptx":
                    // PowerPoint
                    System.Diagnostics.Process.Start("POWERPNT.EXE", '"' + clsCommonSettting.strHMDocumentFilePath + '"');
                    break;

                default:
                    // その他の形式
                    System.Diagnostics.Process.Start('"' + clsCommonSettting.strHMDocumentFilePath+ '"');
                    break;
            }
        }
        #endregion

        #region メソッド
        /// <summary>
        /// 大中小分類名コンボボックス更新処理
        /// </summary>
        /// <param name="strLargeClassName"></param>
        /// <param name="strMiddleClassName"></param>
        /// <param name="strSmallClassName"></param>
        private void ComboBoxClassSetting(string strLargeClassName, string strMiddleClassName, string strSmallClassName)
        {
            // 大分類名コンボボックス設定
            if (!string.IsNullOrEmpty(strLargeClassName))
            {
                // 大分類設定
                comboBoxLargeClass.Text = strLargeClassName;

                // 中分類初期化
                if (string.IsNullOrEmpty(strMiddleClassName))
                {
                    string[] strMiddleClassList = clsModule.getMiddleClassNameList(strLargeClassName);
                    comboBoxMiddleClass.Items.Clear();
                    foreach (string strMiddleClass in strMiddleClassList)
                    {
                        comboBoxMiddleClass.Items.Add(strMiddleClass.Trim());
                    }
                    if (comboBoxMiddleClass.Items.Count > 0) comboBoxMiddleClass.SelectedIndex = 0;
                }

                // 小分類初期化
                if (string.IsNullOrEmpty(strSmallClassName))
                {
                    string[] strSmallClassList = clsModule.getSmallClassNameList(strLargeClassName, comboBoxMiddleClass.Text);
                    comboBoxSmallClass.Items.Clear();
                    foreach (string strSmallClass in strSmallClassList)
                    {
                        comboBoxSmallClass.Items.Add(strSmallClass.Trim());
                    }
                    if (comboBoxSmallClass.Items.Count > 0) comboBoxSmallClass.SelectedIndex = 0;
                }
            }

            // 中分類名コンボボックス設定
            if (!string.IsNullOrEmpty(strMiddleClassName))
            {
                // 中分類設定
                comboBoxMiddleClass.Text = strMiddleClassName;

                // 小分類初期化
                if (string.IsNullOrEmpty(strSmallClassName))
                {
                    string[] strSmallClassList = clsModule.getSmallClassNameList(comboBoxLargeClass.Text, strMiddleClassName);
                    comboBoxSmallClass.Items.Clear();
                    foreach (string strSmallClass in strSmallClassList)
                    {
                        comboBoxSmallClass.Items.Add(strSmallClass.Trim());
                    }
                    if (comboBoxSmallClass.Items.Count > 0) comboBoxSmallClass.SelectedIndex = 0;
                }
            }


            // 小分類名コンボボックス設定
            if (!string.IsNullOrEmpty(strSmallClassName))
            {
                // 小分類設定
                comboBoxSmallClass.Text = strSmallClassName;
            }

            // 文書の名称（サンプル）設定
            string[] strDocumentExampleList = clsModule.getDocumentExampleList(comboBoxLargeClass.Text, comboBoxMiddleClass.Text, comboBoxSmallClass.Text);
            listBoxDocumentExample.Items.Clear();
            foreach (string strDocumentExample in strDocumentExampleList)
            {
                listBoxDocumentExample.Items.Add(strDocumentExample);
            }
            if (listBoxDocumentExample.Items.Count <= 0) listBoxDocumentExample.Items.Add("");
            listBoxDocumentExample.SelectedIndex = 0;

            // 登録内容更新
            string[] strLargeCodeStrit = comboBoxLargeClass.Text.Trim().Split('.');
            string strLargeCode = (strLargeCodeStrit.Count() > 1) ? strLargeCodeStrit[0] : "Z";
            string[] strMiddleCodeStrit = comboBoxMiddleClass.Text.Trim().Split('.');
            string strMiddleCode = (strMiddleCodeStrit.Count() > 1) ? strMiddleCodeStrit[0] : "99";
            string[] strSmallCodeStrit = comboBoxSmallClass.Text.Trim().Split(' ');
            string strSmallCode = (strSmallCodeStrit.Count() > 1) ? strSmallCodeStrit[0] : "9999";
            labelClassNo.Text = string.Format("{0}.{1}.{2}", strLargeCode, strMiddleCode, strSmallCode);
            labelSmallClass.Text = (string.IsNullOrEmpty(comboBoxSmallClass.Text.Trim())) ? "" : comboBoxSmallClass.Text.Substring(5);

            // SAB秘設定変更時処理
            SABSetting();

        }

        /// <summary>
        /// SAB秘設定変更時処理
        /// </summary>
        private void SABSetting()
        {
            // SABデータ取得処理
            IEnumerable<ClassAttributeSetting.AttributeSetting> searchData;
            searchData = clsModule.getAttributeSettingList(comboBoxLargeClass.Text, comboBoxMiddleClass.Text, comboBoxSmallClass.Text, listBoxDocumentExample.SelectedItem.ToString());

            // 検索結果表示
            foreach (ClassAttributeSetting.AttributeSetting row in searchData)
            {
                // SAB秘表示
                labelSecrecyLevel.Text = row.SecrecyLevel;
                break;
            }

        }

        /// <summary>
        /// ユーザー設定読み込み
        /// </summary>
        private Boolean UserSettingRead()
        {
            Boolean bResult = false;
            try
            {
                //// ユーザー設定ファイルパス作成
                //string strUserSettingFilePath = GetFileSystemPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SETFILENAME;

                //// ユーザー設定ファイルが存在しない場合はデフォルト設定を書き込む
                //if (!File.Exists(strUserSettingFilePath))
                //{
                //    UserSettingWrite();
                //}

                // ユーザー設定ファイルパス作成
                // 20171009 修正
                string strUserSettingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SETFOLDERNAME + "\\" + USER_SETFILENAME;

                string oldStrUserSettingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SET_OLDFOLDERNAME + "\\" + USER_SETFILENAME;

                // 
                //// ユーザー設定ファイルが存在しない場合はデフォルト設定を書き込む
                //if (!File.Exists(strUserSettingFilePath))
                //{
                //    UserSettingWrite();
                //}

                // 20171009 修正
                // ユーザー設定ファイルが存在しない場合
                if (!File.Exists(strUserSettingFilePath))
                {
                    if (!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SET_OLDFOLDERNAME + "\\" + USER_SETFILENAME))
                    {
                        // 旧フォルダにもファイルが無い場合はデフォルト設定を書き込む
                        UserSettingWrite();
                    }
                    else
                    {
                        // 20171011 コメントアウト
                        //// 旧フォルダにある場合は旧ファイルを読み込む
                        //strUserSettingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SET_OLDFOLDERNAME + "\\" + USER_SETFILENAME;
                        // 20171011 追加（旧フォルダから新フォルダに設定ファイルコピー）
                        try
                        {
                            File.Copy(oldStrUserSettingFilePath , strUserSettingFilePath);
                        }
                        catch
                        {

                        }
                    }
                }

                //XmlSerializerオブジェクトの作成
                System.Xml.Serialization.XmlSerializer serXmlUserRead = new System.Xml.Serialization.XmlSerializer(typeof(UserSettings));

                //ファイルを開く
                System.IO.StreamReader stmUserReader = new System.IO.StreamReader(strUserSettingFilePath, Encoding.GetEncoding("shift_jis"));

                //XMLファイルから読み込み、逆シリアル化する
                clsUserSettting = (UserSettings)serXmlUserRead.Deserialize(stmUserReader);

                //閉じる
                stmUserReader.Close();

                bResult = true;
            }
            catch(Exception e)
            {
            }
            return bResult;
        }

        /// <summary>
        /// ユーザー設定書き込み
        /// </summary>
        public Boolean UserSettingWrite()
        {
            Boolean bResult = false;
            try
            {
                //// ユーザー設定ファイルパス作成
                //string strUserSettingFilePath = GetFileSystemPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SETFILENAME;

                // ユーザー設定ファイルパス作成
                // 20171009 修正
                //string strUserSettingFilePath = GetFileSystemPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SETFILENAME;
                string strUserSettingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + USER_SETFOLDERNAME;
                if (!Directory.Exists(strUserSettingFilePath))
                {
                    // フォルダ作成
                    System.IO.Directory.CreateDirectory(strUserSettingFilePath);
                }
                strUserSettingFilePath += "\\" + USER_SETFILENAME;

                //// 20171011 追加（設定ファイルが存在している場合は、以降の処理不要）
                //if (File.Exists(strUserSettingFilePath)) return true;

                //XmlSerializerオブジェクトを作成
                System.Xml.Serialization.XmlSerializer serXmlUserWrite = new System.Xml.Serialization.XmlSerializer(typeof(UserSettings));

                //ファイルを開く
                System.IO.StreamWriter stmUserWrite = new System.IO.StreamWriter(strUserSettingFilePath, false, Encoding.GetEncoding("shift_jis"));

                //シリアル化し、XMLファイルに保存する
                serXmlUserWrite.Serialize(stmUserWrite, clsUserSettting);

                //閉じる
                stmUserWrite.Close();

                bResult = true;
            }
            catch
            {
            }
            return bResult;
        }

        /// <summary>
        /// 共通設定読み込み
        /// </summary>
        private Boolean CommonSettingRead()
        {
            Boolean bResult = false;
            try
            {
                // 20171009 修正（格納フォルダ先を変更）
                // 共通設定ファイルパス作成
                //string strCommonSettingFilePath = GetFileSystemPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + COMMON_SETFILENAME;
                  string strCommonSettingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + COMMON_SETFOLDERNAME + "\\" + COMMON_SETFILENAME;

                //// 共通設定ファイルパス作成
                //string strCommonSettingFilePath = GetFileSystemPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + COMMON_SETFILENAME;

                // 共通設定ファイルが存在しない場合はデフォルト設定を書き込む
                if(!File.Exists(strCommonSettingFilePath))
                {
                    CommonSettingWrite();
                }

                //XmlSerializerオブジェクトの作成
                System.Xml.Serialization.XmlSerializer serXmlCommonRead = new System.Xml.Serialization.XmlSerializer(typeof(CommonSettings));

                //ファイルを開く
                System.IO.StreamReader stmCommonReader = new System.IO.StreamReader(strCommonSettingFilePath, Encoding.GetEncoding("shift_jis"));

                //XMLファイルから読み込み、逆シリアル化する
                clsCommonSettting = (CommonSettings)serXmlCommonRead.Deserialize(stmCommonReader);

                //閉じる
                stmCommonReader.Close();

                bResult = true;
            }
            catch
            {
            }
            return bResult;
        }

        /// <summary>
        /// 共通設定設定書き込み
        /// </summary>
        private Boolean CommonSettingWrite()
        {
            Boolean bResult = false;
            try
            {
                //// 共通設定ファイルパス作成
                //string strCommonSettingFilePath = GetFileSystemPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + COMMON_SETFILENAME;

                // 20170714 修正（格納フォルダ先を変更）
                // 共通設定ファイルパス作成
                //string strCommonSettingFilePath = GetFileSystemPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + COMMON_SETFILENAME;
                string strCommonSettingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + COMMON_SETFOLDERNAME;
                if (!Directory.Exists(strCommonSettingFilePath))
                {
                    // フォルダ作成
                    System.IO.Directory.CreateDirectory(strCommonSettingFilePath);
                }
                strCommonSettingFilePath += "\\" + COMMON_SETFILENAME;

                //XmlSerializerオブジェクトの作成
                System.Xml.Serialization.XmlSerializer serXmlCommonWrite = new System.Xml.Serialization.XmlSerializer(typeof(CommonSettings));

                //ファイルを開く
                System.IO.StreamWriter stmCommonWrite = new System.IO.StreamWriter(strCommonSettingFilePath, false, Encoding.GetEncoding("shift_jis"));

                //シリアル化し、XMLファイルに保存する
                serXmlCommonWrite.Serialize(stmCommonWrite, clsCommonSettting);

                //閉じる
                stmCommonWrite.Close();

                bResult = true;
            }
            catch
            {
            }
            return bResult;
        }

        /// <summary>
        /// 設定ファイルパス取得
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        private static string GetFileSystemPath(Environment.SpecialFolder folder)
        {
            // パスを取得
            string path = String.Format(@"{0}\{1}\{2}",
                Environment.GetFolderPath(folder),  // ベース・パス
                Application.CompanyName,            // 会社名
                Application.ProductName);           // 製品名

            // パスのフォルダを作成
            lock (typeof(Application))
            {
                if (Directory.Exists(path) == false)
                {
                    Directory.CreateDirectory(path);
                }
            }
            return path;
        }

        #endregion
    }
}
