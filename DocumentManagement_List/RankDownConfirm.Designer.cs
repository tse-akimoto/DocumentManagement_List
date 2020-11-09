namespace DocumentManagement_List
{
    partial class RankDownConfirm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RankDownConfirm));
            this.lbl_RankDownMsg1 = new System.Windows.Forms.Label();
            this.btn_RankDownOK = new System.Windows.Forms.Button();
            this.chk_SkipMsg = new System.Windows.Forms.CheckBox();
            this.lbl_RankFrom = new System.Windows.Forms.Label();
            this.lbl_RankMsgFrom = new System.Windows.Forms.Label();
            this.lbl_RankTo = new System.Windows.Forms.Label();
            this.lbl_RankDownMsg2 = new System.Windows.Forms.Label();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_Skip = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl_RankDownMsg1
            // 
            resources.ApplyResources(this.lbl_RankDownMsg1, "lbl_RankDownMsg1");
            this.lbl_RankDownMsg1.Name = "lbl_RankDownMsg1";
            // 
            // btn_RankDownOK
            // 
            resources.ApplyResources(this.btn_RankDownOK, "btn_RankDownOK");
            this.btn_RankDownOK.Name = "btn_RankDownOK";
            this.btn_RankDownOK.UseVisualStyleBackColor = true;
            this.btn_RankDownOK.Click += new System.EventHandler(this.btn_RankDownOK_Click);
            // 
            // chk_SkipMsg
            // 
            resources.ApplyResources(this.chk_SkipMsg, "chk_SkipMsg");
            this.chk_SkipMsg.Name = "chk_SkipMsg";
            this.chk_SkipMsg.UseVisualStyleBackColor = true;
            // 
            // lbl_RankFrom
            // 
            resources.ApplyResources(this.lbl_RankFrom, "lbl_RankFrom");
            this.lbl_RankFrom.Name = "lbl_RankFrom";
            // 
            // lbl_RankMsgFrom
            // 
            resources.ApplyResources(this.lbl_RankMsgFrom, "lbl_RankMsgFrom");
            this.lbl_RankMsgFrom.Name = "lbl_RankMsgFrom";
            // 
            // lbl_RankTo
            // 
            resources.ApplyResources(this.lbl_RankTo, "lbl_RankTo");
            this.lbl_RankTo.Name = "lbl_RankTo";
            // 
            // lbl_RankDownMsg2
            // 
            resources.ApplyResources(this.lbl_RankDownMsg2, "lbl_RankDownMsg2");
            this.lbl_RankDownMsg2.Name = "lbl_RankDownMsg2";
            // 
            // btn_Cancel
            // 
            resources.ApplyResources(this.btn_Cancel, "btn_Cancel");
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // btn_Skip
            // 
            resources.ApplyResources(this.btn_Skip, "btn_Skip");
            this.btn_Skip.Name = "btn_Skip";
            this.btn_Skip.UseVisualStyleBackColor = true;
            this.btn_Skip.Click += new System.EventHandler(this.btn_Skip_Click);
            // 
            // RankDownConfirm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btn_Skip);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.lbl_RankDownMsg2);
            this.Controls.Add(this.lbl_RankTo);
            this.Controls.Add(this.lbl_RankMsgFrom);
            this.Controls.Add(this.lbl_RankFrom);
            this.Controls.Add(this.chk_SkipMsg);
            this.Controls.Add(this.btn_RankDownOK);
            this.Controls.Add(this.lbl_RankDownMsg1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RankDownConfirm";
            this.Load += new System.EventHandler(this.RankDownConfirm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_RankDownMsg1;
        private System.Windows.Forms.Button btn_RankDownOK;
        private System.Windows.Forms.CheckBox chk_SkipMsg;
        private System.Windows.Forms.Label lbl_RankFrom;
        private System.Windows.Forms.Label lbl_RankMsgFrom;
        private System.Windows.Forms.Label lbl_RankTo;
        private System.Windows.Forms.Label lbl_RankDownMsg2;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Skip;
    }
}