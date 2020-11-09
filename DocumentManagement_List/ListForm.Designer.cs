namespace DocumentManagement_List
{
    partial class ListForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ListForm));
            this.label1 = new System.Windows.Forms.Label();
            this.buttonStart = new System.Windows.Forms.Button();
            this.panelInput = new System.Windows.Forms.Panel();
            this.groupBoxFile = new System.Windows.Forms.GroupBox();
            this.checkBoxZipTarget = new System.Windows.Forms.CheckBox();
            this.buttonReference = new System.Windows.Forms.Button();
            this.buttonSearch = new System.Windows.Forms.Button();
            this.textBoxFolderPath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.radioButtonSubFolderInc = new System.Windows.Forms.RadioButton();
            this.radioButtonSubFolderNotInc = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBoxUpdate = new System.Windows.Forms.CheckBox();
            this.lblUpdate = new System.Windows.Forms.Label();
            this.lblCreated = new System.Windows.Forms.Label();
            this.checkBoxSAB_Other = new System.Windows.Forms.CheckBox();
            this.buttonRefine = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.panelFileType = new System.Windows.Forms.Panel();
            this.radioButtonTypeAll = new System.Windows.Forms.RadioButton();
            this.radioButtonTypeOfficePdf = new System.Windows.Forms.RadioButton();
            this.radioButtonTypeOffice = new System.Windows.Forms.RadioButton();
            this.radioButtonTypePdf = new System.Windows.Forms.RadioButton();
            this.checkBoxCreate = new System.Windows.Forms.CheckBox();
            this.dateTimePickerUpdateTo = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerUpdateFrom = new System.Windows.Forms.DateTimePicker();
            this.label10 = new System.Windows.Forms.Label();
            this.dateTimePickerCreateTo = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerCreateFrom = new System.Windows.Forms.DateTimePicker();
            this.label11 = new System.Windows.Forms.Label();
            this.textBoxClassNo = new System.Windows.Forms.TextBox();
            this.textBoxFileName = new System.Windows.Forms.TextBox();
            this.checkBoxSAB_None = new System.Windows.Forms.CheckBox();
            this.checkBoxSAB_B = new System.Windows.Forms.CheckBox();
            this.checkBoxSAB_A = new System.Windows.Forms.CheckBox();
            this.checkBoxSAB_S = new System.Windows.Forms.CheckBox();
            this.checkBoxSAB_All = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBoxAll = new System.Windows.Forms.CheckBox();
            this.panelOutput = new System.Windows.Forms.Panel();
            this.radioButtonDesignation = new System.Windows.Forms.RadioButton();
            this.radioButtonBatch = new System.Windows.Forms.RadioButton();
            this.dataGridViewList = new System.Windows.Forms.DataGridView();
            this.ZipFormat = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FileType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.作成日 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.更新日 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.選択 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.文書分類 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.機密区分 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnShelfLife = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnSaveLives = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Creator = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnLastModifiedBy = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.機密区分_Hidden = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnZipFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnZipCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.削除ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ファイルを開くToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip_OpenFolder = new System.Windows.Forms.ToolStripMenuItem();
            this.buttonExcelOutput = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.buttonStop = new System.Windows.Forms.Button();
            this.panelProcess = new System.Windows.Forms.Panel();
            this.panelInput.SuspendLayout();
            this.groupBoxFile.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panelFileType.SuspendLayout();
            this.panelOutput.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewList)).BeginInit();
            this.contextMenuStrip.SuspendLayout();
            this.panelProcess.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // buttonStart
            // 
            resources.ApplyResources(this.buttonStart, "buttonStart");
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // panelInput
            // 
            resources.ApplyResources(this.panelInput, "panelInput");
            this.panelInput.BackColor = System.Drawing.Color.White;
            this.panelInput.Controls.Add(this.groupBoxFile);
            this.panelInput.Controls.Add(this.groupBox2);
            this.panelInput.Name = "panelInput";
            // 
            // groupBoxFile
            // 
            resources.ApplyResources(this.groupBoxFile, "groupBoxFile");
            this.groupBoxFile.Controls.Add(this.checkBoxZipTarget);
            this.groupBoxFile.Controls.Add(this.buttonReference);
            this.groupBoxFile.Controls.Add(this.buttonSearch);
            this.groupBoxFile.Controls.Add(this.textBoxFolderPath);
            this.groupBoxFile.Controls.Add(this.label2);
            this.groupBoxFile.Controls.Add(this.radioButtonSubFolderInc);
            this.groupBoxFile.Controls.Add(this.radioButtonSubFolderNotInc);
            this.groupBoxFile.Name = "groupBoxFile";
            this.groupBoxFile.TabStop = false;
            // 
            // checkBoxZipTarget
            // 
            resources.ApplyResources(this.checkBoxZipTarget, "checkBoxZipTarget");
            this.checkBoxZipTarget.Name = "checkBoxZipTarget";
            this.checkBoxZipTarget.UseVisualStyleBackColor = true;
            // 
            // buttonReference
            // 
            resources.ApplyResources(this.buttonReference, "buttonReference");
            this.buttonReference.Name = "buttonReference";
            this.buttonReference.UseVisualStyleBackColor = true;
            this.buttonReference.Click += new System.EventHandler(this.buttonReference_Click);
            // 
            // buttonSearch
            // 
            resources.ApplyResources(this.buttonSearch, "buttonSearch");
            this.buttonSearch.Name = "buttonSearch";
            this.buttonSearch.UseVisualStyleBackColor = true;
            this.buttonSearch.Click += new System.EventHandler(this.buttonSearch_Click);
            // 
            // textBoxFolderPath
            // 
            resources.ApplyResources(this.textBoxFolderPath, "textBoxFolderPath");
            this.textBoxFolderPath.Name = "textBoxFolderPath";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // radioButtonSubFolderInc
            // 
            resources.ApplyResources(this.radioButtonSubFolderInc, "radioButtonSubFolderInc");
            this.radioButtonSubFolderInc.Checked = true;
            this.radioButtonSubFolderInc.Name = "radioButtonSubFolderInc";
            this.radioButtonSubFolderInc.TabStop = true;
            this.radioButtonSubFolderInc.UseVisualStyleBackColor = true;
            // 
            // radioButtonSubFolderNotInc
            // 
            resources.ApplyResources(this.radioButtonSubFolderNotInc, "radioButtonSubFolderNotInc");
            this.radioButtonSubFolderNotInc.Name = "radioButtonSubFolderNotInc";
            this.radioButtonSubFolderNotInc.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Controls.Add(this.checkBoxUpdate);
            this.groupBox2.Controls.Add(this.lblUpdate);
            this.groupBox2.Controls.Add(this.lblCreated);
            this.groupBox2.Controls.Add(this.checkBoxSAB_Other);
            this.groupBox2.Controls.Add(this.buttonRefine);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.panelFileType);
            this.groupBox2.Controls.Add(this.checkBoxCreate);
            this.groupBox2.Controls.Add(this.dateTimePickerUpdateTo);
            this.groupBox2.Controls.Add(this.dateTimePickerUpdateFrom);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.dateTimePickerCreateTo);
            this.groupBox2.Controls.Add(this.dateTimePickerCreateFrom);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.textBoxClassNo);
            this.groupBox2.Controls.Add(this.textBoxFileName);
            this.groupBox2.Controls.Add(this.checkBoxSAB_None);
            this.groupBox2.Controls.Add(this.checkBoxSAB_B);
            this.groupBox2.Controls.Add(this.checkBoxSAB_A);
            this.groupBox2.Controls.Add(this.checkBoxSAB_S);
            this.groupBox2.Controls.Add(this.checkBoxSAB_All);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // checkBoxUpdate
            // 
            resources.ApplyResources(this.checkBoxUpdate, "checkBoxUpdate");
            this.checkBoxUpdate.Name = "checkBoxUpdate";
            this.checkBoxUpdate.UseVisualStyleBackColor = true;
            this.checkBoxUpdate.CheckedChanged += new System.EventHandler(this.checkBoxUpdate_CheckedChanged);
            // 
            // lblUpdate
            // 
            resources.ApplyResources(this.lblUpdate, "lblUpdate");
            this.lblUpdate.Name = "lblUpdate";
            // 
            // lblCreated
            // 
            resources.ApplyResources(this.lblCreated, "lblCreated");
            this.lblCreated.Name = "lblCreated";
            // 
            // checkBoxSAB_Other
            // 
            resources.ApplyResources(this.checkBoxSAB_Other, "checkBoxSAB_Other");
            this.checkBoxSAB_Other.Checked = true;
            this.checkBoxSAB_Other.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSAB_Other.Name = "checkBoxSAB_Other";
            this.checkBoxSAB_Other.UseVisualStyleBackColor = true;
            // 
            // buttonRefine
            // 
            resources.ApplyResources(this.buttonRefine, "buttonRefine");
            this.buttonRefine.Name = "buttonRefine";
            this.buttonRefine.UseVisualStyleBackColor = true;
            this.buttonRefine.Click += new System.EventHandler(this.buttonRefine_Click);
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // panelFileType
            // 
            this.panelFileType.Controls.Add(this.radioButtonTypeAll);
            this.panelFileType.Controls.Add(this.radioButtonTypeOfficePdf);
            this.panelFileType.Controls.Add(this.radioButtonTypeOffice);
            this.panelFileType.Controls.Add(this.radioButtonTypePdf);
            resources.ApplyResources(this.panelFileType, "panelFileType");
            this.panelFileType.Name = "panelFileType";
            // 
            // radioButtonTypeAll
            // 
            resources.ApplyResources(this.radioButtonTypeAll, "radioButtonTypeAll");
            this.radioButtonTypeAll.Checked = true;
            this.radioButtonTypeAll.Name = "radioButtonTypeAll";
            this.radioButtonTypeAll.TabStop = true;
            this.radioButtonTypeAll.UseVisualStyleBackColor = true;
            // 
            // radioButtonTypeOfficePdf
            // 
            resources.ApplyResources(this.radioButtonTypeOfficePdf, "radioButtonTypeOfficePdf");
            this.radioButtonTypeOfficePdf.Name = "radioButtonTypeOfficePdf";
            this.radioButtonTypeOfficePdf.UseVisualStyleBackColor = true;
            // 
            // radioButtonTypeOffice
            // 
            resources.ApplyResources(this.radioButtonTypeOffice, "radioButtonTypeOffice");
            this.radioButtonTypeOffice.Name = "radioButtonTypeOffice";
            this.radioButtonTypeOffice.UseVisualStyleBackColor = true;
            // 
            // radioButtonTypePdf
            // 
            resources.ApplyResources(this.radioButtonTypePdf, "radioButtonTypePdf");
            this.radioButtonTypePdf.Name = "radioButtonTypePdf";
            this.radioButtonTypePdf.UseVisualStyleBackColor = true;
            // 
            // checkBoxCreate
            // 
            resources.ApplyResources(this.checkBoxCreate, "checkBoxCreate");
            this.checkBoxCreate.Name = "checkBoxCreate";
            this.checkBoxCreate.UseVisualStyleBackColor = true;
            this.checkBoxCreate.CheckedChanged += new System.EventHandler(this.checkBoxCreate_CheckedChanged);
            // 
            // dateTimePickerUpdateTo
            // 
            resources.ApplyResources(this.dateTimePickerUpdateTo, "dateTimePickerUpdateTo");
            this.dateTimePickerUpdateTo.Name = "dateTimePickerUpdateTo";
            // 
            // dateTimePickerUpdateFrom
            // 
            resources.ApplyResources(this.dateTimePickerUpdateFrom, "dateTimePickerUpdateFrom");
            this.dateTimePickerUpdateFrom.Name = "dateTimePickerUpdateFrom";
            // 
            // label10
            // 
            resources.ApplyResources(this.label10, "label10");
            this.label10.Name = "label10";
            // 
            // dateTimePickerCreateTo
            // 
            resources.ApplyResources(this.dateTimePickerCreateTo, "dateTimePickerCreateTo");
            this.dateTimePickerCreateTo.Name = "dateTimePickerCreateTo";
            // 
            // dateTimePickerCreateFrom
            // 
            resources.ApplyResources(this.dateTimePickerCreateFrom, "dateTimePickerCreateFrom");
            this.dateTimePickerCreateFrom.Name = "dateTimePickerCreateFrom";
            // 
            // label11
            // 
            resources.ApplyResources(this.label11, "label11");
            this.label11.Name = "label11";
            // 
            // textBoxClassNo
            // 
            resources.ApplyResources(this.textBoxClassNo, "textBoxClassNo");
            this.textBoxClassNo.Name = "textBoxClassNo";
            // 
            // textBoxFileName
            // 
            resources.ApplyResources(this.textBoxFileName, "textBoxFileName");
            this.textBoxFileName.Name = "textBoxFileName";
            // 
            // checkBoxSAB_None
            // 
            resources.ApplyResources(this.checkBoxSAB_None, "checkBoxSAB_None");
            this.checkBoxSAB_None.Checked = true;
            this.checkBoxSAB_None.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSAB_None.Name = "checkBoxSAB_None";
            this.checkBoxSAB_None.UseVisualStyleBackColor = true;
            this.checkBoxSAB_None.Click += new System.EventHandler(this.checkBoxSAB_CheckedChanged);
            // 
            // checkBoxSAB_B
            // 
            resources.ApplyResources(this.checkBoxSAB_B, "checkBoxSAB_B");
            this.checkBoxSAB_B.Checked = true;
            this.checkBoxSAB_B.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSAB_B.Name = "checkBoxSAB_B";
            this.checkBoxSAB_B.UseVisualStyleBackColor = true;
            this.checkBoxSAB_B.Click += new System.EventHandler(this.checkBoxSAB_CheckedChanged);
            // 
            // checkBoxSAB_A
            // 
            resources.ApplyResources(this.checkBoxSAB_A, "checkBoxSAB_A");
            this.checkBoxSAB_A.Checked = true;
            this.checkBoxSAB_A.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSAB_A.Name = "checkBoxSAB_A";
            this.checkBoxSAB_A.UseVisualStyleBackColor = true;
            this.checkBoxSAB_A.Click += new System.EventHandler(this.checkBoxSAB_CheckedChanged);
            // 
            // checkBoxSAB_S
            // 
            resources.ApplyResources(this.checkBoxSAB_S, "checkBoxSAB_S");
            this.checkBoxSAB_S.Checked = true;
            this.checkBoxSAB_S.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSAB_S.Name = "checkBoxSAB_S";
            this.checkBoxSAB_S.UseVisualStyleBackColor = true;
            this.checkBoxSAB_S.Click += new System.EventHandler(this.checkBoxSAB_CheckedChanged);
            // 
            // checkBoxSAB_All
            // 
            resources.ApplyResources(this.checkBoxSAB_All, "checkBoxSAB_All");
            this.checkBoxSAB_All.Checked = true;
            this.checkBoxSAB_All.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSAB_All.Name = "checkBoxSAB_All";
            this.checkBoxSAB_All.UseVisualStyleBackColor = true;
            this.checkBoxSAB_All.Click += new System.EventHandler(this.checkBoxSAB_All_CheckedChanged);
            // 
            // label9
            // 
            resources.ApplyResources(this.label9, "label9");
            this.label9.Name = "label9";
            // 
            // label8
            // 
            resources.ApplyResources(this.label8, "label8");
            this.label8.Name = "label8";
            // 
            // label6
            // 
            resources.ApplyResources(this.label6, "label6");
            this.label6.Name = "label6";
            // 
            // label5
            // 
            resources.ApplyResources(this.label5, "label5");
            this.label5.Name = "label5";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // checkBoxAll
            // 
            resources.ApplyResources(this.checkBoxAll, "checkBoxAll");
            this.checkBoxAll.Name = "checkBoxAll";
            this.checkBoxAll.UseVisualStyleBackColor = true;
            this.checkBoxAll.CheckedChanged += new System.EventHandler(this.checkBoxAll_CheckedChanged);
            // 
            // panelOutput
            // 
            resources.ApplyResources(this.panelOutput, "panelOutput");
            this.panelOutput.BackColor = System.Drawing.Color.White;
            this.panelOutput.Controls.Add(this.radioButtonDesignation);
            this.panelOutput.Controls.Add(this.radioButtonBatch);
            this.panelOutput.Name = "panelOutput";
            // 
            // radioButtonDesignation
            // 
            resources.ApplyResources(this.radioButtonDesignation, "radioButtonDesignation");
            this.radioButtonDesignation.Name = "radioButtonDesignation";
            this.radioButtonDesignation.UseVisualStyleBackColor = true;
            // 
            // radioButtonBatch
            // 
            resources.ApplyResources(this.radioButtonBatch, "radioButtonBatch");
            this.radioButtonBatch.Checked = true;
            this.radioButtonBatch.Name = "radioButtonBatch";
            this.radioButtonBatch.TabStop = true;
            this.radioButtonBatch.UseVisualStyleBackColor = true;
            // 
            // dataGridViewList
            // 
            this.dataGridViewList.AllowDrop = true;
            this.dataGridViewList.AllowUserToAddRows = false;
            this.dataGridViewList.AllowUserToDeleteRows = false;
            this.dataGridViewList.AllowUserToResizeRows = false;
            resources.ApplyResources(this.dataGridViewList, "dataGridViewList");
            this.dataGridViewList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ZipFormat,
            this.FileName,
            this.FileType,
            this.作成日,
            this.更新日,
            this.選択,
            this.文書分類,
            this.機密区分,
            this.ColumnFilePath,
            this.ColumnShelfLife,
            this.ColumnSaveLives,
            this.Creator,
            this.ColumnLastModifiedBy,
            this.機密区分_Hidden,
            this.ColumnZipFilePath,
            this.ColumnZipCount});
            this.dataGridViewList.MultiSelect = false;
            this.dataGridViewList.Name = "dataGridViewList";
            this.dataGridViewList.RowHeadersVisible = false;
            this.dataGridViewList.RowTemplate.Height = 21;
            this.dataGridViewList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewList.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewList_CellClick);
            this.dataGridViewList.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridViewList_CellMouseDoubleClick);
            this.dataGridViewList.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridViewList_CellMouseDown);
            this.dataGridViewList.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridViewList_CellPainting);
            this.dataGridViewList.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridViewList_ColumnHeaderMouseClick);
            this.dataGridViewList.DragDrop += new System.Windows.Forms.DragEventHandler(this.dataGridViewList_DragDrop);
            this.dataGridViewList.DragEnter += new System.Windows.Forms.DragEventHandler(this.dataGridViewList_DragEnter);
            this.dataGridViewList.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridViewList_KeyDown);
            // 
            // ZipFormat
            // 
            resources.ApplyResources(this.ZipFormat, "ZipFormat");
            this.ZipFormat.Name = "ZipFormat";
            this.ZipFormat.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // FileName
            // 
            resources.ApplyResources(this.FileName, "FileName");
            this.FileName.Name = "FileName";
            this.FileName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // FileType
            // 
            resources.ApplyResources(this.FileType, "FileType");
            this.FileType.Name = "FileType";
            this.FileType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // 作成日
            // 
            resources.ApplyResources(this.作成日, "作成日");
            this.作成日.Name = "作成日";
            this.作成日.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // 更新日
            // 
            resources.ApplyResources(this.更新日, "更新日");
            this.更新日.Name = "更新日";
            this.更新日.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // 選択
            // 
            this.選択.FalseValue = "0";
            resources.ApplyResources(this.選択, "選択");
            this.選択.Name = "選択";
            this.選択.TrueValue = "1";
            // 
            // 文書分類
            // 
            resources.ApplyResources(this.文書分類, "文書分類");
            this.文書分類.Name = "文書分類";
            this.文書分類.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // 機密区分
            // 
            resources.ApplyResources(this.機密区分, "機密区分");
            this.機密区分.Name = "機密区分";
            this.機密区分.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnFilePath
            // 
            resources.ApplyResources(this.ColumnFilePath, "ColumnFilePath");
            this.ColumnFilePath.Name = "ColumnFilePath";
            this.ColumnFilePath.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnShelfLife
            // 
            resources.ApplyResources(this.ColumnShelfLife, "ColumnShelfLife");
            this.ColumnShelfLife.Name = "ColumnShelfLife";
            this.ColumnShelfLife.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnSaveLives
            // 
            resources.ApplyResources(this.ColumnSaveLives, "ColumnSaveLives");
            this.ColumnSaveLives.Name = "ColumnSaveLives";
            this.ColumnSaveLives.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Creator
            // 
            resources.ApplyResources(this.Creator, "Creator");
            this.Creator.Name = "Creator";
            this.Creator.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnLastModifiedBy
            // 
            resources.ApplyResources(this.ColumnLastModifiedBy, "ColumnLastModifiedBy");
            this.ColumnLastModifiedBy.Name = "ColumnLastModifiedBy";
            this.ColumnLastModifiedBy.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // 機密区分_Hidden
            // 
            resources.ApplyResources(this.機密区分_Hidden, "機密区分_Hidden");
            this.機密区分_Hidden.Name = "機密区分_Hidden";
            this.機密区分_Hidden.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnZipFilePath
            // 
            resources.ApplyResources(this.ColumnZipFilePath, "ColumnZipFilePath");
            this.ColumnZipFilePath.Name = "ColumnZipFilePath";
            this.ColumnZipFilePath.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnZipCount
            // 
            resources.ApplyResources(this.ColumnZipCount, "ColumnZipCount");
            this.ColumnZipCount.Name = "ColumnZipCount";
            this.ColumnZipCount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // contextMenuStrip
            // 
            this.contextMenuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.削除ToolStripMenuItem,
            this.ファイルを開くToolStripMenuItem,
            this.toolStrip_OpenFolder});
            this.contextMenuStrip.Name = "contextMenuStrip1";
            resources.ApplyResources(this.contextMenuStrip, "contextMenuStrip");
            // 
            // 削除ToolStripMenuItem
            // 
            this.削除ToolStripMenuItem.Name = "削除ToolStripMenuItem";
            resources.ApplyResources(this.削除ToolStripMenuItem, "削除ToolStripMenuItem");
            this.削除ToolStripMenuItem.Click += new System.EventHandler(this.削除ToolStripMenuItem_Click);
            // 
            // ファイルを開くToolStripMenuItem
            // 
            this.ファイルを開くToolStripMenuItem.Name = "ファイルを開くToolStripMenuItem";
            resources.ApplyResources(this.ファイルを開くToolStripMenuItem, "ファイルを開くToolStripMenuItem");
            this.ファイルを開くToolStripMenuItem.Click += new System.EventHandler(this.ファイルを開くToolStripMenuItem_Click);
            // 
            // toolStrip_OpenFolder
            // 
            this.toolStrip_OpenFolder.Name = "toolStrip_OpenFolder";
            resources.ApplyResources(this.toolStrip_OpenFolder, "toolStrip_OpenFolder");
            this.toolStrip_OpenFolder.Click += new System.EventHandler(this.toolStrip_OpenFolder_Click);
            // 
            // buttonExcelOutput
            // 
            resources.ApplyResources(this.buttonExcelOutput, "buttonExcelOutput");
            this.buttonExcelOutput.Name = "buttonExcelOutput";
            this.buttonExcelOutput.UseVisualStyleBackColor = true;
            this.buttonExcelOutput.Click += new System.EventHandler(this.buttonExcelOutput_Click);
            // 
            // labelStatus
            // 
            resources.ApplyResources(this.labelStatus, "labelStatus");
            this.labelStatus.Name = "labelStatus";
            // 
            // buttonStop
            // 
            resources.ApplyResources(this.buttonStop, "buttonStop");
            this.buttonStop.Name = "buttonStop";
            this.buttonStop.UseVisualStyleBackColor = true;
            this.buttonStop.Click += new System.EventHandler(this.buttonStop_Click);
            // 
            // panelProcess
            // 
            this.panelProcess.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panelProcess.Controls.Add(this.buttonStop);
            this.panelProcess.Controls.Add(this.labelStatus);
            resources.ApplyResources(this.panelProcess, "panelProcess");
            this.panelProcess.Name = "panelProcess";
            // 
            // ListForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.Controls.Add(this.panelProcess);
            this.Controls.Add(this.buttonExcelOutput);
            this.Controls.Add(this.checkBoxAll);
            this.Controls.Add(this.dataGridViewList);
            this.Controls.Add(this.panelOutput);
            this.Controls.Add(this.panelInput);
            this.Controls.Add(this.buttonStart);
            this.Controls.Add(this.label1);
            this.Name = "ListForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ListForm_FormClosed);
            this.Load += new System.EventHandler(this.ListForm_Load);
            this.panelInput.ResumeLayout(false);
            this.groupBoxFile.ResumeLayout(false);
            this.groupBoxFile.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.panelFileType.ResumeLayout(false);
            this.panelFileType.PerformLayout();
            this.panelOutput.ResumeLayout(false);
            this.panelOutput.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewList)).EndInit();
            this.contextMenuStrip.ResumeLayout(false);
            this.panelProcess.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.Panel panelInput;
        private System.Windows.Forms.Panel panelOutput;
        private System.Windows.Forms.RadioButton radioButtonDesignation;
        private System.Windows.Forms.RadioButton radioButtonBatch;
        private System.Windows.Forms.DataGridView dataGridViewList;
        private System.Windows.Forms.CheckBox checkBoxAll;
        private System.Windows.Forms.Button buttonExcelOutput;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem 削除ToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBoxFile;
        private System.Windows.Forms.Button buttonReference;
        private System.Windows.Forms.Button buttonSearch;
        private System.Windows.Forms.TextBox textBoxFolderPath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton radioButtonSubFolderInc;
        private System.Windows.Forms.RadioButton radioButtonSubFolderNotInc;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panelFileType;
        private System.Windows.Forms.RadioButton radioButtonTypeOffice;
        private System.Windows.Forms.RadioButton radioButtonTypePdf;
        private System.Windows.Forms.CheckBox checkBoxCreate;
        private System.Windows.Forms.DateTimePicker dateTimePickerUpdateTo;
        private System.Windows.Forms.DateTimePicker dateTimePickerUpdateFrom;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.DateTimePicker dateTimePickerCreateTo;
        private System.Windows.Forms.DateTimePicker dateTimePickerCreateFrom;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBoxClassNo;
        private System.Windows.Forms.TextBox textBoxFileName;
        private System.Windows.Forms.CheckBox checkBoxSAB_None;
        private System.Windows.Forms.CheckBox checkBoxSAB_B;
        private System.Windows.Forms.CheckBox checkBoxSAB_A;
        private System.Windows.Forms.CheckBox checkBoxSAB_S;
        private System.Windows.Forms.CheckBox checkBoxSAB_All;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonRefine;
        private System.Windows.Forms.RadioButton radioButtonTypeOfficePdf;
        private System.Windows.Forms.ToolStripMenuItem ファイルを開くToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStrip_OpenFolder;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button buttonStop;
        private System.Windows.Forms.Panel panelProcess;
        private System.Windows.Forms.CheckBox checkBoxZipTarget;
        private System.Windows.Forms.CheckBox checkBoxSAB_Other;
        private System.Windows.Forms.Label lblCreated;
        private System.Windows.Forms.Label lblUpdate;
        private System.Windows.Forms.CheckBox checkBoxUpdate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ZipFormat;
        private System.Windows.Forms.DataGridViewTextBoxColumn FileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn FileType;
        private System.Windows.Forms.DataGridViewTextBoxColumn 作成日;
        private System.Windows.Forms.DataGridViewTextBoxColumn 更新日;
        private System.Windows.Forms.DataGridViewCheckBoxColumn 選択;
        private System.Windows.Forms.DataGridViewTextBoxColumn 文書分類;
        private System.Windows.Forms.DataGridViewTextBoxColumn 機密区分;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnShelfLife;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSaveLives;
        private System.Windows.Forms.DataGridViewTextBoxColumn Creator;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnLastModifiedBy;
        private System.Windows.Forms.DataGridViewTextBoxColumn 機密区分_Hidden;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnZipFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnZipCount;
        private System.Windows.Forms.RadioButton radioButtonTypeAll;
    }
}

