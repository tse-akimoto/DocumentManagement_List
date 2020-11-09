﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Microsoft.Office.Core;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentManagement_List.Properties;   // step2 iwasa

namespace DocumentManagement_List
{
    public partial class SettingForm : Form
    {
        #region <定数定義>

        /// <summary>
        /// SAB秘 S秘
        /// </summary>
        public const string SECRECY_PROPERTY_S    = "SecrecyS";

        /// <summary>
        /// SAB秘 A秘
        /// </summary>
        public const string SECRECY_PROPERTY_A    = "SecrecyA";

        /// <summary>
        /// SAB秘 B秘
        /// </summary>
        public const string SECRECY_PROPERTY_B    = "SecrecyB";

        /// <summary>
        /// SAB秘 以外
        /// </summary>
        public const string SECRECY_PROPERTY_ELSE = "SecrecyNone";

        // 共通設定関連

        /// <summary>
        /// デフォルト機密区分
        /// </summary>
        public const string COMMON_SETDEF_SECLV = SECRECY_PROPERTY_S;

        /// <summary>
        /// 事業所コード
        /// </summary>
        public const string COMMON_SETDEF_OFFICECODE = "HLI";

        /// <summary>
        /// 格納フォルダ名
        /// </summary>
        public const string COMMON_SETFOLDERNAME = "SAB";

        /// <summary>
        /// ファイル名 デフォルト機密区分
        /// </summary>
        public const string COMMON_SETFILENAME   = "common_setting.config";

        // スタンプ関連

        /// <summary>
        /// スタンプ倍率
        /// </summary>
        public const double STAMP_MAGNIFICATION = 1.3331;  // スタンプ倍率

        // Excelに貼り付けられたスタンプを判別するため使用するので、重複しないような文字列にする
        /// <summary>
        /// スタンプを識別するための文字列
        /// </summary>
        protected string STAMP_SHAPE_NAME = "HONDA_SECRECY_STAMP";

        // スタンプ位置

        /// <summary>
        /// スタンプ位置
        /// </summary>
        private const int MARGIN_TOP = 10;

        /// <summary>
        /// スタンプ位置
        /// </summary>
        private const int MARGIN_RIGHT = 10;

        /// <summary>
        /// ファイルパス最大文字数
        /// </summary>
        public const int FILEPATH_OUT_RANGE = 200;

        /// <summary>
        /// Officeプロパティ
        /// </summary>
        public enum Property
        {
            SecrecyLevel = 0, // 機密区分
            ClassNo,          // 文書番号
            OfficeCode,       // 事業所コード
        }

        /// <summary>
        /// SAB機密区分
        /// </summary>
        public enum Secrecy
        {
            S,   // 機密区分
            A,   // 文書番号
            B,   // 事業所コード
            None // 事業所コード
        }
        #endregion


        #region <内部変数>

        /// <summary>
        /// 共通設定クラス
        /// </summary>
        public CommonSettings clsCommonSettting;

        /// <summary>
        /// エラーフラグ
        /// </summary>
        public Boolean commonFileReadCompleted;

        /// <summary>
        /// スタンプ処理済みか
        /// </summary>
        public bool IsStampProc = false;

        /// <summary>
        /// 機密区分必須登録モード切替え
        /// </summary>
        private bool mustRegistMode = false;

        #endregion

        #region <コンストラクタ>
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SettingForm()
        {
            // 共通設定エラー時処理
            commonFileReadCompleted = this.CommonSettingRead();
            if (commonFileReadCompleted == false)
            {
                return;
            }

            // 各コンポーネント初期化
            InitializeComponent();

            // タイトル
            System.Diagnostics.FileVersionInfo ver =
            System.Diagnostics.FileVersionInfo.GetVersionInfo(
            System.Reflection.Assembly.GetExecutingAssembly().Location);

            // タイトル
            string AssemblyName = ver.FileVersion;
            this.Text = this.Text + " " + AssemblyName;

            // 言語設定を表示
            this.lblLanguage.Text = System.Threading.Thread.CurrentThread.CurrentUICulture.ToString();
#if DEBUG
            this.lblLanguage.Visible = true;
#endif
        }
        #endregion


        #region <フォームイベント>
        /// <summary>
        /// フォームロードイベント
        /// </summary>
        public void FormSetting_Load(object sender, EventArgs e)
        {
            string filePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string filePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
            string filePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // ファイルプロパティ情報取得
            filePropertySecrecyLevel = clsCommonSettting.strDefaultSecrecyLevel;
            filePropertyOfficeCode = clsCommonSettting.strOfficeCode;

            // プロパティファイルにSAB機密区分が設定されていない場合は標準値を設定
            if (string.IsNullOrWhiteSpace(filePropertySecrecyLevel))
            {
                filePropertySecrecyLevel = clsCommonSettting.strDefaultSecrecyLevel;
            }

            // ラジオボタンをセット
            this.SetSABRadioButton(filePropertySecrecyLevel);
        }

        /// <summary>
        /// フォームキーダウン処理
        /// </summary>
        protected void FormSetting_KeyDown(object sender, KeyEventArgs e)
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
                btnRegist_Click(sender, e);
            }
        }

        /// <summary>
        /// 後で登録するボタンのクリックイベント
        /// </summary>
        public void buttonNotRegist_Click(object sender, EventArgs e)
        {
            // フォームを閉じる
            this.Close();
        }

        /// <summary>
        /// ラジオボタンを変更したときのイベント
        /// </summary>
        public void btnSAB_CheckedChanged(object sender, EventArgs e)
        {
            // SAB区分のテキスト変更
            Secrecy selectedSecrecy = this.GetSelectedSecrecyLevel();
            this.lblSABSetting.Text = this.GetSABText(selectedSecrecy);

            // SAB区分ラジオボタン 背景色変更
            this.ChangeBackColorSAB(ref this.rdoS);
            this.ChangeBackColorSAB(ref this.rdoA);
            this.ChangeBackColorSAB(ref this.rdoB);
            this.ChangeBackColorSAB(ref this.rdoElse);
        }

        /// <summary>
        /// 閉じるボタンをクリックしたときのイベント
        /// </summary>
        public void btnClose_Click(object sender, EventArgs e)
        {
            // フォームを閉じる
            this.Close();
        }

        /// <summary>
        /// 表示切替ボタン変更時のイベント
        /// </summary>
        public void btnChange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                string textStampON = Properties.Resources.StampDisplay_ON;
                string textStampOFF = Properties.Resources.StampDisplay_OFF;
                lblDisplay.Text = chkChange.Checked == true ? textStampON : textStampOFF;

            }
            catch (Exception ex)
            {
                MessageBox.Show(Properties.Resources.msgViewChangeError,
                    Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
#if DEBUG
            //string textStampON = Properties.Resources.StampDisplay_ON;
            //string textStampOFF = Properties.Resources.StampDisplay_OFF;
            //lblDisplay.Text = chkChange.Checked == true ? textStampON : textStampOFF;
#endif
        }

        /// <summary>
        /// フォームを閉じるときのイベント
        /// </summary>
        public void FormSetting_FormClosing(object sender, FormClosingEventArgs e)
        {
            bool canClose = this.CanClose();
            e.Cancel = canClose;
        }

        /// <summary>
        /// 登録ボタンをクリックしたときのイベント
        /// </summary>
        public void btnRegist_Click(object sender, EventArgs e)
        {
            Console.WriteLine(System.Threading.Thread.CurrentThread.CurrentUICulture);

            string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
            string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // ファイルプロパティセット
            Secrecy selectedSecrecy = this.GetSelectedSecrecyLevel();
            string SABCode = this.GetSABCode(selectedSecrecy);
            clsCommonSettting.strDefaultSecrecyLevel = SABCode;

            // ダイアログを閉じる
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
        #endregion


        #region <メソッド>

        /// <summary>
        /// 共通設定読み込み
        /// </summary>
        /// /// <returns>true:読込み成功、false:読込み失敗</returns>
        private Boolean CommonSettingRead()
        {
            clsCommonSettting = new CommonSettings();

            try
            {
                // 共通設定ファイルパス作成
                string strCommonSettingFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + COMMON_SETFOLDERNAME + "\\" + COMMON_SETFILENAME;

                // 共通設定ファイルが存在しない場合はデフォルト設定を書き込む
                if (File.Exists(strCommonSettingFilePath) == false)
                {
                    this.CommonSettingWrite();
                    return false;
                }

                //XmlSerializerオブジェクトの作成
                System.Xml.Serialization.XmlSerializer serXmlCommonRead = new System.Xml.Serialization.XmlSerializer(typeof(CommonSettings));

                //ファイルを開く
                System.IO.StreamReader stmCommonReader = new System.IO.StreamReader(strCommonSettingFilePath, Encoding.GetEncoding("shift_jis"));

                //XMLファイルから読み込み、逆シリアル化する
                clsCommonSettting = (CommonSettings)serXmlCommonRead.Deserialize(stmCommonReader);

                //閉じる
                stmCommonReader.Close();

                return true;
            }
            catch (Exception ex)
            {
                // 原因不明のエラー
                MessageBox.Show(ex.ToString(),
                     Resources.msgError,
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Hand);
            }

            return false;
        }

        /// <summary>
        /// 共通設定設定書き込み
        /// </summary>
        private Boolean CommonSettingWrite()
        {
            Boolean bResult = false;

            try
            {
                // 共通設定ファイルパス作成
                string strCommonSettingFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    COMMON_SETFOLDERNAME
                    );

                if (!Directory.Exists(strCommonSettingFilePath))
                {
                    // フォルダ作成
                    System.IO.Directory.CreateDirectory(strCommonSettingFilePath);
                }
                strCommonSettingFilePath = Path.Combine(strCommonSettingFilePath,
                        COMMON_SETFILENAME
                        );

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
                // 更新できない場合はエラーを返す
                bResult = false;
            }
            return bResult;
        }

        /// <summary>
        /// 機密区分が登録済みか
        /// </summary>
        /// <returns>true:登録済</returns>
        public bool IsSecrecyInfoRegistered()
        {
            string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 文書番号
            string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // プロパティ情報取得
            strFilePropertySecrecyLevel = clsCommonSettting.strDefaultSecrecyLevel;
            strFilePropertyOfficeCode = clsCommonSettting.strOfficeCode;

            // 機密区分が空白、または事業所コードが自事業所コードではない
            if (string.IsNullOrEmpty(strFilePropertySecrecyLevel) ||
                strFilePropertyOfficeCode != this.clsCommonSettting.strOfficeCode )
            {
                return false;
            }

            // 機密区分のいずれにも該当しない
            if (SECRECY_PROPERTY_S    != strFilePropertySecrecyLevel &&
                SECRECY_PROPERTY_A    != strFilePropertySecrecyLevel &&
                SECRECY_PROPERTY_B    != strFilePropertySecrecyLevel &&
                SECRECY_PROPERTY_ELSE != strFilePropertySecrecyLevel )
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 機密区分必須登録モード切替え
        /// </summary>
        /// <returns>true:登録必須モードON、false:登録必須モードOFF</returns>
        public bool MustRegistMode
        {
            set
            {
                this.mustRegistMode = value;

                // 閉じる機能を切替え
                this.btnClose.Enabled = !value;
                this.ControlBox = !value;
            }
            get
            {
                return this.mustRegistMode;
            }
        }

        /// <summary>
        /// ラジオボタンの選択状態から背景色を更新
        /// </summary>
        /// <param name="radioSAB">対象のラジオボタン</param>
        private void ChangeBackColorSAB(ref RadioButton radioSAB)
        {
            if (radioSAB.Checked == true)
            {
                radioSAB.BackColor = Color.Green;
            }
            else
            {
                radioSAB.BackColor = Color.Gray;
            }
        }

        /// <summary>
        /// SAB機密区分コードからラジオボタンを選択
        /// </summary>
        /// <param name="secrecyCode">SAB機密区分コード</param>
        private void SetSABRadioButton(string secrecyCode)
        {
            if (SECRECY_PROPERTY_S.Equals(secrecyCode))
            {
                this.rdoS.Checked = true;

                return;
            }

            if (SECRECY_PROPERTY_A.Equals(secrecyCode))
            {
                this.rdoA.Checked = true;

                return;
            }

            if (SECRECY_PROPERTY_B.Equals(secrecyCode))
            {
                this.rdoB.Checked = true;

                return;
            }

            this.rdoElse.Checked = true;
        }

        /// <summary>
        /// ラジオボタンの選択状態からSAB機密区分を取得
        /// </summary>
        /// <returns>選択中のSAB機密区分 列挙体</returns>
        private Secrecy GetSelectedSecrecyLevel()
        {
            if (this.rdoS.Checked == true)
            {
                return Secrecy.S;
            }

            if (this.rdoA.Checked == true)
            {
                return Secrecy.A;
            }

            if (this.rdoB.Checked == true)
            {
                return Secrecy.B;
            }

            return Secrecy.None;
        }

        /// <summary>
        /// SAB機密区分列挙体からSAB機密区分のテキストを取得
        /// </summary>
        private string GetSABText(Secrecy secrecy)
        {
            if (secrecy == Secrecy.S)
            {
                return this.rdoS.Text;
            }

            if (secrecy == Secrecy.A)
            {
                return this.rdoA.Text;
            }

            if (secrecy == Secrecy.B)
            {
                return this.rdoB.Text;
            }

            return this.rdoElse.Text; ;
        }

        /// <summary>
        /// ラジオボタンの選択状態からSAB機密区分コードを取得
        /// </summary>
        /// <param name="secrecy"></param>
        /// <returns></returns>
        private string GetSABCode(Secrecy secrecy)
        {
            if (secrecy == Secrecy.S)
            {
                return SECRECY_PROPERTY_S;
            }

            if (secrecy == Secrecy.A)
            {
                return SECRECY_PROPERTY_A;
            }

            if (secrecy == Secrecy.B)
            {
                return SECRECY_PROPERTY_B;
            }

            return SECRECY_PROPERTY_ELSE;
        }

        /// <summary>
        /// SAB機密区分列挙体に対応するスタンプ画像を取得
        /// </summary>
        /// <param name="secrecy">SAB機密区分列挙体</param>
        /// <returns>スタンプ画像</returns>
        protected Bitmap GetStampImage(Secrecy secrecy)
        {
            if (secrecy == Secrecy.S)
            {
                return Properties.Resources.StampS;
            }

            if (secrecy == Secrecy.A)
            {
                return Properties.Resources.StampA;
            }

            if (secrecy == Secrecy.B)
            {
                return Properties.Resources.StampB;
            }

            return null;
        }

        /// <summary>
        /// SAB機密区分の登録状態から閉じれる状態か確認
        /// </summary>
        /// <returns>true:フォーム閉じる不可</returns>
        private bool CanClose()
        {
            bool closingCancel = false;

            // 登録必須の状態ではない場合は閉じて良い
            if (this.MustRegistMode == false)
            {
                return closingCancel;
            }

            // 機密区分が登録されている場合は閉じて良い
            if (this.IsSecrecyInfoRegistered() == true)
            {
                return closingCancel;
            }

            // 閉じて良い条件に該当しなかった場合は閉じない
            closingCancel = true;


            return closingCancel;
        }

        /// <summary>
        /// 画像透過処理
        /// </summary>
        /// <param name="bitmap"></param>
        /// <returns>透過済み画像</returns>
        protected Bitmap CreateAlphaImage(Bitmap bitmap, float alpha)
        {
            int imageWidth = bitmap.Width;
            int imageHeight = bitmap.Height;

            // 新しいビットマップを用意
            Bitmap alphaImage = new Bitmap(imageWidth, imageHeight);

            using (Graphics graphics = Graphics.FromImage(alphaImage))
            {
                // ColorMatrixオブジェクトの作成
                System.Drawing.Imaging.ColorMatrix cm = 
                    new System.Drawing.Imaging.ColorMatrix();

                // ColorMatrixの行列の値を変更して、アルファ値が0.5に変更されるようにする
                cm.Matrix00 = 1;
                cm.Matrix11 = 1;
                cm.Matrix22 = 1;
                cm.Matrix33 = (1f - alpha);
                cm.Matrix44 = 1;

                // ImageAttributesオブジェクトの作成
                System.Drawing.Imaging.ImageAttributes imageAttributes = 
                    new System.Drawing.Imaging.ImageAttributes();

                // ColorMatrixを設定
                imageAttributes.SetColorMatrix(cm);

                // ImageAttributesを使用して画像を描画
                graphics.DrawImage(bitmap,
                            new Rectangle(0, 0, imageWidth, imageHeight),
                            0,
                            0,
                            imageWidth,
                            imageHeight,
                            GraphicsUnit.Pixel,
                            imageAttributes);
            }

            return alphaImage;
        }
        #endregion


        #region <Overrideメソッド>

        /// <summary>
        /// スタンプ貼付け処理
        /// </summary>
        protected virtual Boolean SetStamp(Secrecy secrecyLevel)
        {
            // Excel・Word・PowerPointのスタンプ貼付け処理を子クラスで実装

            return true;
        }

        /// <summary>
        /// ドキュメントのプロパティ取得
        /// </summary>
        /// <param name="strClassNo">文書分類番号</param>
        /// <param name="strSecrecyLevel">機密区分</param>
        /// <param name="bStamp">スタンプ有無</param>off
        public virtual void GetDocumentProperty(ref string strSecrecyLevel, ref string strOfficeCod, ref string strOfficeCode)
        {
            // Excel・Word・PowerPointのプロパティ取得処理を子クラスで実装
        }


        /// <summary>
        /// ドキュメントのプロパティ設定
        /// オーバーライド用のメソッド
        /// </summary>
        /// <param name="strClassNo"></param>
        /// <param name="strSecrecyLevel"></param>
        public virtual bool SetDocumentProperty(string strSecrecyLevel)
        {
            // Excel・Word・PowerPointのプロパティ設定処理を子クラスで実装

            return true;
        }

        #endregion

        #region GCP機密文書(SA秘)一覧表 リンクラベル

        /// <summary>
        /// 機密区分の判別方法ドキュメントの更新処理
        /// </summary>
        public void UpdateDocument()   // step2
        {
            // ネットワークに接続されているか
            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
            {
                string _serverFilePath = clsCommonSettting.strSABListServerPath;
                string _localFilePath = clsCommonSettting.strSABListLocalPath;

                try
                {
                    // 設定値のチェック
                    if (_serverFilePath == "" || _localFilePath == "")
                    {
                        MessageBox.Show(Resources.msgFailedReadDocumentPath, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);  // step2 iwasa
                        return;
                    }

                    // ファイルの有無
                    if (File.Exists(_serverFilePath) == true)
                    {
                        // ローカルの保存先があるか
                        string _localDir = Path.GetDirectoryName(_localFilePath);

                        if (Directory.Exists(_localDir) == false)
                        {
                            // フォルダ作成
                            Directory.CreateDirectory(_localDir);
                        }

                        // サーバとローカルのファイルを比較する
                        DateTime _serverDocument = File.GetLastWriteTime(_serverFilePath);
                        DateTime _localDocument = File.GetLastWriteTime(_localFilePath);

                        if (_serverDocument > _localDocument)
                        {
                            // ローカルファイルを更新する
                            File.Copy(_serverFilePath, _localFilePath, true);
                        }
                    }
                }
                catch
                {
                    // ファイル移動関連でエラーが発生する可能性があるが
                    // 処理として問題ないためスルーする
                }
            }
        }

        /// <summary>
        /// リンクラベルクリック処理
        /// </summary>
        private void lnkLblDoc_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // GCP機密文書(SA秘)一覧表を開く
            string _localFilePath = clsCommonSettting.strSABListLocalPath;

            if (File.Exists(_localFilePath) == true)
            {
                Process.Start(_localFilePath);
            }
            else
            {
                MessageBox.Show(Resources.msgFailedOpenGCPdocument, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);   // step2 iwasa
            }
        }

        #endregion

        #region Officeスタンプ貼り付け

        /// <summary>
        /// 文字列からSecrecyを取得
        /// </summary>
        /// <param name="secrecyCode">機密区分文字列</param>
        /// <returns>機密区分設定値</returns>
        private Secrecy GetSecrecy(string SecrecyLevel)
        {
            Secrecy ret = Secrecy.None;

            // 機密区分設定
            switch (SecrecyLevel)
            {
                case SECRECY_PROPERTY_S:
                    ret = Secrecy.S;
                    break;
                case SECRECY_PROPERTY_A:
                    ret = Secrecy.A;
                    break;
                case SECRECY_PROPERTY_B:
                    ret = Secrecy.B;
                    break;
                default:
                    ret = Secrecy.None;
                    break;
            }

            return ret;
        }

        /// <summary>
        /// スタンプを貼付け
        /// </summary>
        /// <param name="secrecyCode">機密区分文字列</param>
        /// <returns>true:機密区分が異なる false:機密区分変更なし</returns>
        public bool IsChangedSecrecy(string beforeSecrecyLevel)
        {
            bool ret = false;

            Secrecy AfterSecrecy = this.GetSelectedSecrecyLevel();
            Secrecy beforeSecrecy = GetSecrecy(beforeSecrecyLevel);

            if (AfterSecrecy != beforeSecrecy)
            {
                ret = true;
            }

            return ret;
        }

        /// <summary>
        /// スタンプを貼付け
        /// </summary>
        /// <param name="filetype">Office形式</param>
        /// <param name="fileName">ファイル名</param>
        /// <param name="beforeSecrecyLevel">変更後機密区分</param>
        /// <returns>true:成功 false:失敗</returns>
        public bool SetStamp(int filetype, string fileName, string beforeSecrecyLevel)
        {
            // ファイルプロパティセット
            Secrecy selectedSecrecy = this.GetSelectedSecrecyLevel();
            bool IsChanged = IsChangedSecrecy(beforeSecrecyLevel);
            bool setResultIsOK = true;

            if (IsChanged != false)
            {
                // Secrecyに変更があった場合
                switch (filetype)
                {
                    case ListForm.EXTENSION_EXCEL:
                        setResultIsOK = this.SetStampExcel(selectedSecrecy, fileName);
                        break;
                    case ListForm.EXTENSION_WORD:
                        setResultIsOK = this.SetStampWord(selectedSecrecy, fileName);
                        break;
                    case ListForm.EXTENSION_POWERPOINT:
                        setResultIsOK = this.SetStampPowerPoint(selectedSecrecy, fileName);
                        break;
                    default:
                        // その他
                        break;
                }
            }

            return setResultIsOK;
        }

        /// <summary>
        /// スタンプ貼付け処理
        /// </summary>
        /// <param name="secrecyLevel">機密区分</param>
        /// <param name="fileName">ファイル名</param>
        /// <returns>true:成功 false:失敗</returns>
        public Boolean SetStampExcel(Secrecy secrecyLevel , string fileName)
        {
            // 一時ファイル名取得
            string imageFilePath = System.IO.Path.GetTempFileName();
            string excelName = fileName;
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkbooks = null;
            Bitmap bmpSrc = null;
            try
            {
                if (IsFileInUse(excelName) != false)
                {
                    return false;
                }

                excelApp = new Excel.Application();
                excelApp.Visible = false;

                // Excelファイルをオープンする
                excelWorkbooks = (Excel.Workbook)(excelApp.Workbooks.Open(
                  excelName  // オープンするExcelファイル名
                ));

                // スタンプ表示OFF・区分"以外"の場合はスタンプをセットしない
                // スタンプ画像を削除して終了
                if (this.chkChange.Checked == false || this.rdoElse.Checked == true)
                {
                    // 指定した名前のオブジェクトを削除
                    this.DeleteExcelShapes(ref excelWorkbooks, STAMP_SHAPE_NAME);

                    return true;
                }


                // スタンプ画像をリソースから取得
                bmpSrc = this.GetStampImage(secrecyLevel);

                // 画像が取得できない場合は中断
                if (bmpSrc == null) return false;

                // スタンプ倍率変更
                double dStampWidth = bmpSrc.Width / STAMP_MAGNIFICATION;
                double dStampHeight = bmpSrc.Height / STAMP_MAGNIFICATION;

                // 透過処理
                float alpha = (float)(this.nudAlpha.Value * (decimal)0.01);
                bmpSrc = this.CreateAlphaImage(bmpSrc, alpha);

                // ファイルを一時保存
                bmpSrc.Save(imageFilePath, System.Drawing.Imaging.ImageFormat.Png);


                // 指定した名前のオブジェクトを削除
                this.DeleteExcelShapes(ref excelWorkbooks, STAMP_SHAPE_NAME);


                // すべてのExcelシートにスタンプを貼付け
                this.AddStampPicture(ref excelWorkbooks, imageFilePath, (float)dStampWidth, (float)dStampHeight, this.STAMP_SHAPE_NAME);
            }
            catch
            {
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                int FileLength = sjisEnc.GetByteCount(excelName);
                // ファイルパス最大200バイト以上の場合は実行不可とする
                if (FileLength >= FILEPATH_OUT_RANGE)
                {
                    // 文字数エラー
                    string Msg = string.Format(Resources.msgErrorArgumentOutOfRange + "{0}", excelName);

                    MessageBox.Show(Msg, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }

                return false;
            }
            finally
            {
                if (excelWorkbooks != null)
                {
                    excelWorkbooks.Save();
                    excelWorkbooks.Close();
                    excelApp.Quit();

                    // 一時ファイル削除
                    System.IO.File.Delete(imageFilePath);
                }

                if (bmpSrc != null)
                {
                    bmpSrc.Dispose();
                }
            }

            return true;
        }

        /// <summary>
        /// Excelのオブジェクトを削除
        /// </summary>
        /// <param name="excelWorkbooks">対象のExcel</param>
        /// <param name="shapeName">削除するオブジェクト名</param>
        private void DeleteExcelShapes(ref Excel.Workbook excelWorkbooks, string shapeName)
        {
            // すべてのExcelシートからスタンプを削除
            foreach (Excel.Worksheet sheet in excelWorkbooks.Sheets)
            {
                Excel.Shapes excelShapes = (Excel.Shapes)sheet.Shapes;

                // スタンプ画像かオブジェクト名で判定して削除
                foreach (Excel.Shape shape in excelShapes)
                {
                    if (shape.Name == shapeName) shape.Delete();
                }
            }
        }

        /// <summary>
        /// Excelの全てのシートにスタンプを貼付け
        /// </summary>
        /// <param name="ExcelWorkbooks">対象のExcel</param>
        /// <param name="imageFilePath">貼付けるスタンプの画像ファイルパス</param>
        /// <param name="stampWidth">補正する画像の横幅</param>
        /// <param name="stampHeight">補正する画像の縦幅</param>
        private void AddStampPicture(ref Excel.Workbook ExcelWorkbooks, string imageFilePath, float stampWidth, float stampHeight, string stampName)
        {
            foreach (var sheet in ExcelWorkbooks.Sheets)
            {
                Excel.Worksheet workSheet = (Excel.Worksheet)sheet;

                // 画像貼付処理
                Excel.Shapes excelShapes = (Excel.Shapes)workSheet.Shapes;
                Excel.Shape stampShape = excelShapes.AddPicture(imageFilePath,
                                                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                Microsoft.Office.Core.MsoTriState.msoTrue,
                                                                0,
                                                                0,
                                                                (float)stampWidth,
                                                                (float)stampHeight);

                // 貼付けた画像のオブジェクト名を設定
                stampShape.Name = stampName;
            }
        }

        /// <summary>
        /// スタンプ貼付け処理
        /// </summary>
        /// <param name="secrecyLevel">機密区分</param>
        /// <param name="fileName">ファイル名</param>
        /// <returns>true:成功 false:失敗</returns>
        public Boolean SetStampWord(Secrecy secrecyLevel, string fileName)
        {
            // 一時ファイル名取得
            string imageFilePath = System.IO.Path.GetTempFileName();
            Word.Application WordApp = null;
            Word.Document document = null;
            Bitmap bmpSrc = null;
            try
            {
                if (IsFileInUse(fileName) == true)
                {
                    return false;
                }

                // 現在開いているWordの取得
                WordApp = new Word.Application();
                WordApp.Visible = false;

                WordApp.Documents.Open(
                      fileName
                    );

                document = WordApp.ActiveDocument;

                // スタンプ表示OFF・区分"以外"の場合はスタンプをセットしない
                // スタンプ画像を削除して終了
                if (this.chkChange.Checked == false || this.rdoElse.Checked == true)
                {
                    // 指定した名前のオブジェクトを削除
                    this.DeleteWordShapes(ref document, STAMP_SHAPE_NAME);

                    return true;
                }


                // スタンプ画像をリソースから取得
                bmpSrc = this.GetStampImage(secrecyLevel);

                // 画像が取得できない場合は中断
                if (bmpSrc == null) return false;

                // スタンプ倍率変更
                double dStampWidth = bmpSrc.Width / STAMP_MAGNIFICATION;
                double dStampHeight = bmpSrc.Height / STAMP_MAGNIFICATION;

                // 透過処理
                float alpha = (float)(this.nudAlpha.Value * (decimal)0.01);
                bmpSrc = this.CreateAlphaImage(bmpSrc, alpha);

                // ファイルを一時保存
                bmpSrc.Save(imageFilePath, System.Drawing.Imaging.ImageFormat.Png);


                // 指定した名前のオブジェクトを削除
                this.DeleteWordShapes(ref document, STAMP_SHAPE_NAME);


                // スタンプの右上位置を算出
                float topLocation = (float)0 - document.PageSetup.TopMargin + MARGIN_TOP;
                float leftLocation = document.PageSetup.PageWidth - document.PageSetup.RightMargin - (float)dStampWidth - MARGIN_RIGHT;

                // 画像貼付処理
                Word.Shape stampShape = document.Shapes.AddPicture(imageFilePath,
                                                                   Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                   Microsoft.Office.Core.MsoTriState.msoTrue,
                                                                      leftLocation,
                                                                   topLocation,
                                                                   dStampWidth,
                                                                   dStampHeight,
                                                                   document.Range(System.Type.Missing,
                                                                   System.Type.Missing));
                // 貼付けた画像のオブジェクト名を設定
                stampShape.Name = this.STAMP_SHAPE_NAME;
            }
            catch
            {
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                int FileLength = sjisEnc.GetByteCount(fileName);
                // ファイルパス最大200バイト以上の場合は実行不可とする
                if (FileLength >= FILEPATH_OUT_RANGE)
                {
                    // 文字数エラー
                    string Msg = string.Format(Resources.msgErrorArgumentOutOfRange + "{0}", fileName);
                    MessageBox.Show(Msg, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
                return false;
            }
            finally
            {
                if (WordApp != null)
                {
                    WordApp.ActiveDocument.Save();
                    WordApp.ActiveDocument.Close();
                    WordApp.Quit();

                    // 一時ファイル削除
                    System.IO.File.Delete(imageFilePath);
                }

                if (bmpSrc != null)
                {
                    bmpSrc.Dispose();
                }
            }

            return true;
        }

        /// <summary>
        /// Wordのオブジェクトを削除
        /// </summary>
        /// <param name="powerPoint">対象のWord</param>
        /// <param name="shapeName">オブジェクト名</param>
        private void DeleteWordShapes(ref Word.Document wordDocument, string shapeName)
        {
            // スタンプ画像かオブジェクト名で判定して削除
            foreach (Word.Shape shape in wordDocument.Shapes)
            {
                if (shape.Name == shapeName) shape.Delete();
            }
        }

        /// <summary>
        /// スタンプ貼付け処理
        /// </summary>
        /// <param name="secrecyLevel">機密区分</param>
        /// <param name="fileName">ファイル名</param>
        /// <returns>true:成功 false:失敗</returns>
        public Boolean SetStampPowerPoint(Secrecy secrecyLevel, string fileName)
        {
            // 一時ファイル名取得
            string imageFilePath = System.IO.Path.GetTempFileName();
            PowerPoint.Application pptApp = null;
            PowerPoint.Presentations ppPress = null;
            PowerPoint.Presentation pptFile = null;
            PowerPoint._Slide slide = null;
            Bitmap bmpSrc = null;
            try
            {
                // 現在開いているPowerPointの取得
                pptApp = new PowerPoint.Application();
                ppPress = pptApp.Presentations;

                foreach (PowerPoint.Presentation ppt in ppPress)
                {
                    if (ppt.FullName == fileName)
                    {
                        // 既に開いている場合
                        return false;
                    }
                }

                pptFile = ppPress.Open(
                    fileName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse
                );

                if (pptFile == null)
                {
                    return false;
                }

                // 先頭のスライド取得
                slide = (PowerPoint.Slide)pptFile.Slides[1];

                // スタンプ表示OFF・区分"以外"の場合はスタンプをセットしない
                // スタンプ画像を削除して終了
                if (this.chkChange.Checked == false || this.rdoElse.Checked == true)
                {
                    // 指定した名前のオブジェクトを削除
                    this.DeletePowerPointShapes(ref pptApp, ref pptFile, this.STAMP_SHAPE_NAME);

                    return true;
                }


                // スタンプ画像をリソースから取得
                bmpSrc = this.GetStampImage(secrecyLevel);

                // 画像が取得できない場合は中断
                if (bmpSrc == null) return false;

                // スタンプ倍率変更
                double dStampWidth = bmpSrc.Width / STAMP_MAGNIFICATION;
                double dStampHeight = bmpSrc.Height / STAMP_MAGNIFICATION;

                // 透過処理
                float alpha = (float)(this.nudAlpha.Value * (decimal)0.01);
                bmpSrc = this.CreateAlphaImage(bmpSrc, alpha);

                // ファイルを一時保存
                bmpSrc.Save(imageFilePath, System.Drawing.Imaging.ImageFormat.Png);

                // スタンプの水平位置を PPTの幅 - 画像の幅 で算出
                float leftLocation = slide.Master.Width - (float)dStampWidth;


                // 指定した名前のオブジェクトを削除
                this.DeletePowerPointShapes(ref pptApp, ref pptFile, this.STAMP_SHAPE_NAME);


                // 画像貼付処理
                PowerPoint.Shape stampShape = slide.Shapes.AddPicture(imageFilePath,
                                                                      MsoTriState.msoFalse,
                                                                      MsoTriState.msoTrue,
                                                                      leftLocation,
                                                                      0,
                                                                      (float)dStampWidth,
                                                                      (float)dStampHeight);
                // 貼付けた画像のオブジェクト名を設定
                stampShape.Name = this.STAMP_SHAPE_NAME;

                if (pptFile != null)
                {
                    pptFile.Save();
                }
            }
            catch
            {
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                int FileLength = sjisEnc.GetByteCount(fileName);
                // ファイルパス最大200バイト以上の場合は実行不可とする
                if (FileLength >= FILEPATH_OUT_RANGE)
                {
                    // 文字数エラー
                    string Msg = string.Format(Resources.msgErrorArgumentOutOfRange + "{0}", fileName);
                    MessageBox.Show(Msg, Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
                return false;
            }
            finally
            {
                if (pptFile != null)
                {
                    pptFile.Close();
                }

                slide = null;
                pptFile = null;
                ppPress = null;
                pptApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                // 一時ファイル削除
                System.IO.File.Delete(imageFilePath);

                if (bmpSrc != null)
                {
                    bmpSrc.Dispose();
                }
            }

            return true;
        }

        /// <summary>
        /// PowerPointのオブジェクトを削除
        /// </summary>
        /// <param name="powerPoint">対象のPowerPoint</param>
        /// <param name="shapeName">オブジェクト名</param>
        private void DeletePowerPointShapes(ref PowerPoint.Application powerPoint,ref PowerPoint.Presentation Press, string shapeName)
        {
            // すべてのスライドからスタンプを削除
            foreach (PowerPoint._Slide slide in Press.Slides)
            {
                PowerPoint.Shapes shapes = (PowerPoint.Shapes)slide.Shapes;

                // スタンプ画像かオブジェクト名で判定して削除
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (shape.Name == shapeName) shape.Delete();
                }
            }
        }

        /// <summary>
        /// 開いているファイルか？
        /// </summary>
        /// <param name="secrecyCode">対象ファイルパス</param>
        /// <returns>true:開いている false:閉じている</returns>
        private bool IsFileInUse(string path)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read)) { }
            }
            catch (IOException)
            {
                // Exceptionが発生する時は対象ファイルがロックされていることを返す
                return true;
            }

            return false;
        }
        #endregion
    }
}