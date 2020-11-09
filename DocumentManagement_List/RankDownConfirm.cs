using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentManagement_List
{
    public partial class RankDownConfirm : Form
    {
        /// <summary>
        /// スキップフラグ
        /// </summary>
        public bool bMsgSkip = false;

        /// <summary>
        /// 全てフラグ
        /// </summary>
        private bool bBatch = false;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public RankDownConfirm(string rank_from, string rank_to, bool bBatch)
        {
            InitializeComponent();

            this.lbl_RankFrom.Text = rank_from;
            this.lbl_RankTo.Text = rank_to;
            this.bBatch = bBatch;
        }

        /// <summary>
        /// フォームロード
        /// </summary>
        private void RankDownConfirm_Load(object sender, EventArgs e)
        {
            // キャンセルボタンにフォーカスを充てる
            this.ActiveControl = this.btn_Cancel;

            if (!this.bBatch)
            {
                // 一件ずつ処理するモードなのでスキップメッセージを非表示に
                this.chk_SkipMsg.Checked = false;
                // 20170913修正（グレーアウトで表示する）
                //this.chk_SkipMsg.Visible = false;
                this.chk_SkipMsg.Enabled = false;
            }
        }

        /// <summary>
        /// OKボタンクリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_RankDownOK_Click(object sender, EventArgs e)
        {
            this.bMsgSkip = chk_SkipMsg.Checked;

            this.DialogResult = DialogResult.OK;
            // ダイアログを閉じる
            this.Close();
        }

        /// <summary>
        /// スキップボタンクリック
        /// </summary>
        private void btn_Skip_Click(object sender, EventArgs e)
        {
            this.bMsgSkip = chk_SkipMsg.Checked;
            this.DialogResult = DialogResult.Ignore;
            // ダイアログを閉じる
            this.Close();
        }

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.bMsgSkip = chk_SkipMsg.Checked;
            this.DialogResult = DialogResult.Cancel;
            // ダイアログを閉じる
            this.Close();
        }
    }
}
