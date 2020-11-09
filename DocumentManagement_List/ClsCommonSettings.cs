using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentManagement_List
{
    /// <summary>
    /// 共通設定ファイル取得用のクラス
    /// </summary>
    public class CommonSettings
    {
        #region <クラス項目定義>

        /// <summary>
        /// デフォルト機密区分
        /// </summary>
        public string strDefaultSecrecyLevel;

        /// <summary>
        /// 事業所コード
        /// </summary>
        public string strOfficeCode;

        /// <summary>
        /// 言語設定
        /// </summary>
        public string strCulture;

        /// <summary>
        /// 文書のローカルパス
        /// </summary>
        public string strSABListLocalPath;

        /// <summary>
        /// 文書のサーバーパス
        /// </summary>
        public string strSABListServerPath;

        /// <summary>
        /// zip一時解凍先
        /// </summary>
        public string strTempPath;

        /// <summary>
        /// 保存可能リスト
        /// </summary>
        public List<string> lstSecureFolder;

        /// <summary>
        /// 「最終版」を表す文字列
        /// </summary>
        public List<string> lstFinal;

        /// <summary>
        /// 関連会社ファイルパス(サーバー)
        /// </summary>
        public string strGroupDomainListServerPath { get; set; }

        /// <summary>
        /// 関連会社ファイルパス(ローカル)
        /// </summary>
        public string strGroupDomainListLocalPath { get; set; }

        /// <summary>
        /// 役職者ファイルパス(サーバー)
        /// </summary>
        public string strManagerListServerPath { get; set; }

        /// <summary>
        /// 役職者ファイルパス(ローカル)
        /// </summary>
        public string strManagerListLocalPath { get; set; }

        /// <summary>
        /// zip化強制レベル
        /// </summary>
        public int zipLevel { get; set; }

        #endregion

        #region <定数定義>

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 機密区分
        /// </summary>
        public const string COMMON_SETDEF_SECLV = "SecrecyS";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 事業所コード
        /// </summary>
        public const string COMMON_SETDEF_OFFICECODE = "HLI";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 文書のサーバーパス
        /// </summary>
        public const string COMMON_SETDEF_SABLISTSERVERPATH = "C:\\SAB_TEST_SRV\\GCPList.xlsx";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 文書のローカルパス
        /// </summary>
        public const string COMMON_SETDEF_SABLISTLOCALPATH = "C:\\SAB_TEMP\\GCPList.xlsx";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// zip一時解凍先
        /// </summary>
        public const string COMMON_SETDEF_TEMPPATH = "C:\\SAB_TEMP\\WORK";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// セキュアフォルダリスト
        /// </summary>
        public const string COMMON_SETDEF_SECUREFOLDER_1 = "\\" + "\\SRV-FS001";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// セキュアフォルダリスト
        /// </summary>
        public const string COMMON_SETDEF_SECUREFOLDER_2 = "C:\\SAB_TEST_SRV";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 「最終版」を表す文字列
        /// </summary>
        public const string COMMON_SETDEF_FINAL_1 = "最終版";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 「最終版」を表す文字列
        /// </summary>
        public const string COMMON_SETDEF_FINAL_2 = "Final";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 関連会社ファイルパス(サーバー)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTSERVERPATH = "C:\\SAB_TEST_SRV\\DomainList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 関連会社ファイルパス(ローカル)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTLOCALPATH = "C:\\SAB_TEMP\\DomainList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 役職者ファイルパス(サーバー)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTSERVERPATH = "C:\\SAB_TEST_SRV\\ManagerList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 役職者ファイルパス(ローカル)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTLOCALPATH = "C:\\SAB_TEMP\\ManagerList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// zip化強制レベル
        /// </summary>
        public const int COMMON_SETDEF_ZIPLEVEL = 1;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public CommonSettings()
        {
            // デフォルト事業所コード初期化
            strOfficeCode = COMMON_SETDEF_OFFICECODE;

            // デフォルト機密区分初期化
            strDefaultSecrecyLevel = COMMON_SETDEF_SECLV;

            // デフォルト言語コード初期化
            strCulture = System.Threading.Thread.CurrentThread.CurrentUICulture.ToString();

            // デフォルト文書のサーバーパス初期化
            strSABListServerPath = COMMON_SETDEF_SABLISTSERVERPATH;

            // デフォルト文書のローカルパス初期化
            strSABListLocalPath = COMMON_SETDEF_SABLISTLOCALPATH;

            // デフォルトzip一時解凍先初期化
            strTempPath = COMMON_SETDEF_TEMPPATH;

            // デフォルトセキュアフォルダリスト初期化
            lstSecureFolder = new List<string>() { COMMON_SETDEF_SECUREFOLDER_1, COMMON_SETDEF_SECUREFOLDER_2 };

            // デフォルト「最終版」を表す文字列初期化
            lstFinal = new List<string>() { COMMON_SETDEF_FINAL_1, COMMON_SETDEF_FINAL_2 };

            // デフォルト関連会社ファイルパス(サーバー)初期化
            strGroupDomainListServerPath = COMMON_SETDEF_GROUPDOMAINLISTSERVERPATH;

            // デフォルト関連会社ファイルパス(ローカル)初期化
            strGroupDomainListLocalPath = COMMON_SETDEF_GROUPDOMAINLISTLOCALPATH;

            // デフォルト役職者ファイルパス(サーバー)初期化
            strManagerListServerPath = COMMON_SETDEF_MANAGERLISTSERVERPATH;

            // デフォルト役職者ファイルパス(ローカル)初期化
            strManagerListLocalPath = COMMON_SETDEF_MANAGERLISTLOCALPATH;

            // デフォルトzip化強制レベル初期化
            zipLevel = COMMON_SETDEF_ZIPLEVEL;
        }

        #endregion

    }
}
