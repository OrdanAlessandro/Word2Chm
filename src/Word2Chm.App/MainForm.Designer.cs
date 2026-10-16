namespace Word2Chm.App;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            components?.Dispose();
        }
        base.Dispose(disposing);
    }

	#region Windows Form Designer generated code

	private void InitializeComponent()
	{
		_root = new TableLayoutPanel();
		_docxPathLabel = new Label();
		_docxPath = new TextBox();
		_browseDocxButton = new Button();
		_outputDirectoryLabel = new Label();
		_outputDirectory = new TextBox();
		_browseOutputButton = new Button();
		_baseNameLabel = new Label();
		_baseName = new TextBox();
		_contextIdHeaderPathLabel = new Label();
		_contextIdHeaderPath = new TextBox();
		_browseHeaderButton = new Button();
		_chmFileNameLabel = new Label();
		_chmFileName = new TextBox();
		_chmCopyPathLabel = new Label();
		_chmCopyPath = new TextBox();
		_browseChmCopyButton = new Button();
		_pageLevelLabel = new Label();
		_pageLevel = new NumericUpDown();
		_bodyFontSizeLabel = new Label();
		_bodyFontSize = new NumericUpDown();
		_templateDirectoryLabel = new Label();
		_templateDirectory = new TextBox();
		_browseTemplateButton = new Button();
		_footerLabel = new Label();
		_footer = new TextBox();
		_hhcPathLabel = new Label();
		_hhcPath = new TextBox();
		_detectHhcButton = new Button();
		_optionsPanel = new FlowLayoutPanel();
		_compileChm = new CheckBox();
		_openOutput = new CheckBox();
		_openChm = new CheckBox();
		_convertButton = new Button();
		_logLabel = new Label();
		_logPanel = new Panel();
		_log = new TextBox();
		_progress = new ProgressBar();
		_toolStrip = new ToolStrip();
		_newProjectButton = new ToolStripButton();
		_openProjectButton = new ToolStripButton();
		_saveProjectButton = new ToolStripButton();
		_saveProjectAsButton = new ToolStripButton();
		_root.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)_pageLevel).BeginInit();
		((System.ComponentModel.ISupportInitialize)_bodyFontSize).BeginInit();
		_optionsPanel.SuspendLayout();
		_logPanel.SuspendLayout();
		_toolStrip.SuspendLayout();
		SuspendLayout();
		// 
		// _root
		// 
		_root.ColumnCount = 3;
		_root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
		_root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
		_root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
		_root.Controls.Add(_docxPathLabel, 0, 0);
		_root.Controls.Add(_docxPath, 1, 0);
		_root.Controls.Add(_browseDocxButton, 2, 0);
		_root.Controls.Add(_outputDirectoryLabel, 0, 1);
		_root.Controls.Add(_outputDirectory, 1, 1);
		_root.Controls.Add(_browseOutputButton, 2, 1);
		_root.Controls.Add(_baseNameLabel, 0, 2);
		_root.Controls.Add(_baseName, 1, 2);
		_root.Controls.Add(_contextIdHeaderPathLabel, 0, 3);
		_root.Controls.Add(_contextIdHeaderPath, 1, 3);
		_root.Controls.Add(_browseHeaderButton, 2, 3);
		_root.Controls.Add(_chmFileNameLabel, 0, 4);
		_root.Controls.Add(_chmFileName, 1, 4);
		_root.Controls.Add(_chmCopyPathLabel, 0, 5);
		_root.Controls.Add(_chmCopyPath, 1, 5);
		_root.Controls.Add(_browseChmCopyButton, 2, 5);
		_root.Controls.Add(_pageLevelLabel, 0, 6);
		_root.Controls.Add(_pageLevel, 1, 6);
		_root.Controls.Add(_bodyFontSizeLabel, 0, 7);
		_root.Controls.Add(_bodyFontSize, 1, 7);
		_root.Controls.Add(_templateDirectoryLabel, 0, 8);
		_root.Controls.Add(_templateDirectory, 1, 8);
		_root.Controls.Add(_browseTemplateButton, 2, 8);
		_root.Controls.Add(_footerLabel, 0, 9);
		_root.Controls.Add(_footer, 1, 9);
		_root.Controls.Add(_hhcPathLabel, 0, 10);
		_root.Controls.Add(_hhcPath, 1, 10);
		_root.Controls.Add(_detectHhcButton, 2, 10);
		_root.Controls.Add(_optionsPanel, 1, 11);
		_root.Controls.Add(_convertButton, 1, 12);
		_root.Controls.Add(_logLabel, 0, 13);
		_root.Controls.Add(_logPanel, 1, 13);
		_root.Dock = DockStyle.Fill;
		_root.Location = new Point(0, 30);
		_root.Name = "_root";
		_root.Padding = new Padding(12);
		_root.RowCount = 14;
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle());
		_root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
		_root.Size = new Size(900, 631);
		_root.TabIndex = 1;
		// 
		// _docxPathLabel
		// 
		_docxPathLabel.Anchor = AnchorStyles.Left;
		_docxPathLabel.AutoSize = true;
		_docxPathLabel.Location = new Point(15, 21);
		_docxPathLabel.Name = "_docxPathLabel";
		_docxPathLabel.Size = new Size(105, 15);
		_docxPathLabel.TabIndex = 0;
		_docxPathLabel.Text = "Documento Word:";
		_docxPathLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _docxPath
		// 
		_docxPath.Dock = DockStyle.Fill;
		_docxPath.Location = new Point(165, 15);
		_docxPath.Name = "_docxPath";
		_docxPath.Size = new Size(620, 23);
		_docxPath.TabIndex = 0;
		// 
		// _browseDocxButton
		// 
		_browseDocxButton.AutoSize = true;
		_browseDocxButton.Location = new Point(791, 15);
		_browseDocxButton.Name = "_browseDocxButton";
		_browseDocxButton.Size = new Size(94, 27);
		_browseDocxButton.TabIndex = 1;
		_browseDocxButton.Text = "Sfoglia...";
		_browseDocxButton.UseVisualStyleBackColor = true;
		_browseDocxButton.Click += BrowseDocx_Click;
		// 
		// _outputDirectoryLabel
		// 
		_outputDirectoryLabel.Anchor = AnchorStyles.Left;
		_outputDirectoryLabel.AutoSize = true;
		_outputDirectoryLabel.Location = new Point(15, 54);
		_outputDirectoryLabel.Name = "_outputDirectoryLabel";
		_outputDirectoryLabel.Size = new Size(102, 15);
		_outputDirectoryLabel.TabIndex = 2;
		_outputDirectoryLabel.Text = "Cartella di output:";
		_outputDirectoryLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _outputDirectory
		// 
		_outputDirectory.Dock = DockStyle.Fill;
		_outputDirectory.Location = new Point(165, 48);
		_outputDirectory.Name = "_outputDirectory";
		_outputDirectory.Size = new Size(620, 23);
		_outputDirectory.TabIndex = 2;
		// 
		// _browseOutputButton
		// 
		_browseOutputButton.AutoSize = true;
		_browseOutputButton.Location = new Point(791, 48);
		_browseOutputButton.Name = "_browseOutputButton";
		_browseOutputButton.Size = new Size(94, 27);
		_browseOutputButton.TabIndex = 3;
		_browseOutputButton.Text = "Sfoglia...";
		_browseOutputButton.UseVisualStyleBackColor = true;
		_browseOutputButton.Click += BrowseOutput_Click;
		// 
		// _baseNameLabel
		// 
		_baseNameLabel.Anchor = AnchorStyles.Left;
		_baseNameLabel.AutoSize = true;
		_baseNameLabel.Location = new Point(15, 85);
		_baseNameLabel.Name = "_baseNameLabel";
		_baseNameLabel.Size = new Size(70, 15);
		_baseNameLabel.TabIndex = 4;
		_baseNameLabel.Text = "Nome base:";
		_baseNameLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _baseName
		// 
		_baseName.Dock = DockStyle.Fill;
		_baseName.Location = new Point(165, 81);
		_baseName.Name = "_baseName";
		_baseName.Size = new Size(620, 23);
		_baseName.TabIndex = 4;
		// 
		// _contextIdHeaderPathLabel
		// 
		_contextIdHeaderPathLabel.Anchor = AnchorStyles.Left;
		_contextIdHeaderPathLabel.AutoSize = true;
		_contextIdHeaderPathLabel.Location = new Point(15, 116);
		_contextIdHeaderPathLabel.Name = "_contextIdHeaderPathLabel";
		_contextIdHeaderPathLabel.Size = new Size(63, 15);
		_contextIdHeaderPathLabel.TabIndex = 5;
		_contextIdHeaderPathLabel.Text = "File ID (.h):";
		_contextIdHeaderPathLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _contextIdHeaderPath
		// 
		_contextIdHeaderPath.Dock = DockStyle.Fill;
		_contextIdHeaderPath.Location = new Point(165, 110);
		_contextIdHeaderPath.Name = "_contextIdHeaderPath";
		_contextIdHeaderPath.Size = new Size(620, 23);
		_contextIdHeaderPath.TabIndex = 5;
		// 
		// _browseHeaderButton
		// 
		_browseHeaderButton.AutoSize = true;
		_browseHeaderButton.Location = new Point(791, 110);
		_browseHeaderButton.Name = "_browseHeaderButton";
		_browseHeaderButton.Size = new Size(94, 27);
		_browseHeaderButton.TabIndex = 6;
		_browseHeaderButton.Text = "Sfoglia...";
		_browseHeaderButton.UseVisualStyleBackColor = true;
		_browseHeaderButton.Click += BrowseHeader_Click;
		// 
		// _chmFileNameLabel
		// 
		_chmFileNameLabel.Anchor = AnchorStyles.Left;
		_chmFileNameLabel.AutoSize = true;
		_chmFileNameLabel.Location = new Point(15, 147);
		_chmFileNameLabel.Name = "_chmFileNameLabel";
		_chmFileNameLabel.Size = new Size(59, 15);
		_chmFileNameLabel.TabIndex = 7;
		_chmFileNameLabel.Text = "File CHM:";
		_chmFileNameLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _chmFileName
		// 
		_root.SetColumnSpan(_chmFileName, 2);
		_chmFileName.Dock = DockStyle.Fill;
		_chmFileName.Location = new Point(165, 143);
		_chmFileName.Name = "_chmFileName";
		_chmFileName.Size = new Size(720, 23);
		_chmFileName.TabIndex = 7;
		// 
		// _chmCopyPathLabel
		// 
		_chmCopyPathLabel.Anchor = AnchorStyles.Left;
		_chmCopyPathLabel.AutoSize = true;
		_chmCopyPathLabel.Location = new Point(15, 178);
		_chmCopyPathLabel.Name = "_chmCopyPathLabel";
		_chmCopyPathLabel.Size = new Size(93, 15);
		_chmCopyPathLabel.TabIndex = 8;
		_chmCopyPathLabel.Text = "Copia il .chm in:";
		_chmCopyPathLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _chmCopyPath
		// 
		_chmCopyPath.Dock = DockStyle.Fill;
		_chmCopyPath.Location = new Point(165, 172);
		_chmCopyPath.Name = "_chmCopyPath";
		_chmCopyPath.Size = new Size(620, 23);
		_chmCopyPath.TabIndex = 8;
		// 
		// _browseChmCopyButton
		// 
		_browseChmCopyButton.AutoSize = true;
		_browseChmCopyButton.Location = new Point(791, 172);
		_browseChmCopyButton.Name = "_browseChmCopyButton";
		_browseChmCopyButton.Size = new Size(94, 27);
		_browseChmCopyButton.TabIndex = 9;
		_browseChmCopyButton.Text = "Sfoglia...";
		_browseChmCopyButton.UseVisualStyleBackColor = true;
		_browseChmCopyButton.Click += BrowseChmCopy_Click;
		// 
		// _pageLevelLabel
		// 
		_pageLevelLabel.Anchor = AnchorStyles.Left;
		_pageLevelLabel.AutoSize = true;
		_pageLevelLabel.Location = new Point(15, 209);
		_pageLevelLabel.Name = "_pageLevelLabel";
		_pageLevelLabel.Size = new Size(130, 15);
		_pageLevelLabel.TabIndex = 10;
		_pageLevelLabel.Text = "Nuova pagina al livello:";
		_pageLevelLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _pageLevel
		// 
		_pageLevel.Dock = DockStyle.Left;
		_pageLevel.Location = new Point(165, 205);
		_pageLevel.Maximum = new decimal(new int[] { 6, 0, 0, 0 });
		_pageLevel.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
		_pageLevel.Name = "_pageLevel";
		_pageLevel.Size = new Size(120, 23);
		_pageLevel.TabIndex = 10;
		_pageLevel.Value = new decimal(new int[] { 6, 0, 0, 0 });
		// 
		// _bodyFontSizeLabel
		// 
		_bodyFontSizeLabel.Anchor = AnchorStyles.Left;
		_bodyFontSizeLabel.AutoSize = true;
		_bodyFontSizeLabel.Location = new Point(15, 238);
		_bodyFontSizeLabel.Name = "_bodyFontSizeLabel";
		_bodyFontSizeLabel.Size = new Size(124, 15);
		_bodyFontSizeLabel.TabIndex = 11;
		_bodyFontSizeLabel.Text = "Dimensione testo (pt):";
		_bodyFontSizeLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _bodyFontSize
		// 
		_bodyFontSize.DecimalPlaces = 1;
		_bodyFontSize.Dock = DockStyle.Left;
		_bodyFontSize.Increment = new decimal(new int[] { 5, 0, 0, 65536 });
		_bodyFontSize.Location = new Point(165, 234);
		_bodyFontSize.Maximum = new decimal(new int[] { 24, 0, 0, 0 });
		_bodyFontSize.Minimum = new decimal(new int[] { 6, 0, 0, 0 });
		_bodyFontSize.Name = "_bodyFontSize";
		_bodyFontSize.Size = new Size(120, 23);
		_bodyFontSize.TabIndex = 11;
		_bodyFontSize.Value = new decimal(new int[] { 105, 0, 0, 65536 });
		// 
		// _templateDirectoryLabel
		// 
		_templateDirectoryLabel.Anchor = AnchorStyles.Left;
		_templateDirectoryLabel.AutoSize = true;
		_templateDirectoryLabel.Location = new Point(15, 269);
		_templateDirectoryLabel.Name = "_templateDirectoryLabel";
		_templateDirectoryLabel.Size = new Size(91, 15);
		_templateDirectoryLabel.TabIndex = 12;
		_templateDirectoryLabel.Text = "Template (skin):";
		_templateDirectoryLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _templateDirectory
		// 
		_templateDirectory.Dock = DockStyle.Fill;
		_templateDirectory.Location = new Point(165, 263);
		_templateDirectory.Name = "_templateDirectory";
		_templateDirectory.Size = new Size(620, 23);
		_templateDirectory.TabIndex = 12;
		// 
		// _browseTemplateButton
		// 
		_browseTemplateButton.AutoSize = true;
		_browseTemplateButton.Location = new Point(791, 263);
		_browseTemplateButton.Name = "_browseTemplateButton";
		_browseTemplateButton.Size = new Size(94, 27);
		_browseTemplateButton.TabIndex = 13;
		_browseTemplateButton.Text = "Sfoglia...";
		_browseTemplateButton.UseVisualStyleBackColor = true;
		_browseTemplateButton.Click += BrowseTemplate_Click;
		// 
		// _footerLabel
		// 
		_footerLabel.Anchor = AnchorStyles.Left;
		_footerLabel.AutoSize = true;
		_footerLabel.Location = new Point(15, 300);
		_footerLabel.Name = "_footerLabel";
		_footerLabel.Size = new Size(78, 15);
		_footerLabel.TabIndex = 14;
		_footerLabel.Text = "Piè di pagina:";
		_footerLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _footer
		// 
		_root.SetColumnSpan(_footer, 2);
		_footer.Dock = DockStyle.Fill;
		_footer.Location = new Point(165, 296);
		_footer.Name = "_footer";
		_footer.Size = new Size(720, 23);
		_footer.TabIndex = 14;
		// 
		// _hhcPathLabel
		// 
		_hhcPathLabel.Anchor = AnchorStyles.Left;
		_hhcPathLabel.AutoSize = true;
		_hhcPathLabel.Location = new Point(15, 331);
		_hhcPathLabel.Name = "_hhcPathLabel";
		_hhcPathLabel.Size = new Size(50, 15);
		_hhcPathLabel.TabIndex = 15;
		_hhcPathLabel.Text = "hhc.exe:";
		_hhcPathLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _hhcPath
		// 
		_hhcPath.Dock = DockStyle.Fill;
		_hhcPath.Location = new Point(165, 325);
		_hhcPath.Name = "_hhcPath";
		_hhcPath.Size = new Size(620, 23);
		_hhcPath.TabIndex = 15;
		// 
		// _detectHhcButton
		// 
		_detectHhcButton.AutoSize = true;
		_detectHhcButton.Location = new Point(791, 325);
		_detectHhcButton.Name = "_detectHhcButton";
		_detectHhcButton.Size = new Size(94, 27);
		_detectHhcButton.TabIndex = 16;
		_detectHhcButton.Text = "Rileva";
		_detectHhcButton.UseVisualStyleBackColor = true;
		_detectHhcButton.Click += DetectHhc_Click;
		// 
		// _optionsPanel
		// 
		_optionsPanel.AutoSize = true;
		_root.SetColumnSpan(_optionsPanel, 2);
		_optionsPanel.Controls.Add(_compileChm);
		_optionsPanel.Controls.Add(_openChm);
		_optionsPanel.Controls.Add(_openOutput);
		_optionsPanel.Dock = DockStyle.Fill;
		_optionsPanel.Location = new Point(165, 358);
		_optionsPanel.Name = "_optionsPanel";
		_optionsPanel.Size = new Size(720, 25);
		_optionsPanel.TabIndex = 17;
		// 
		// _compileChm
		// 
		_compileChm.AutoSize = true;
		_compileChm.Checked = true;
		_compileChm.CheckState = CheckState.Checked;
		_compileChm.Location = new Point(3, 3);
		_compileChm.Name = "_compileChm";
		_compileChm.Size = new Size(110, 19);
		_compileChm.TabIndex = 0;
		_compileChm.Text = "Compila il .chm";
		_compileChm.UseVisualStyleBackColor = true;
		// 
		// _openOutput
		// 
		_openOutput.AutoSize = true;
		_openOutput.Location = new Point(276, 3);
		_openOutput.Name = "_openOutput";
		_openOutput.Size = new Size(157, 19);
		_openOutput.TabIndex = 2;
		_openOutput.Text = "Apri la cartella al termine";
		_openOutput.UseVisualStyleBackColor = true;
		//
		// _openChm
		//
		_openChm.AutoSize = true;
		_openChm.Location = new Point(119, 3);
		_openChm.Name = "_openChm";
		_openChm.Size = new Size(151, 19);
		_openChm.TabIndex = 1;
		_openChm.Text = "Apri il CHM al termine";
		_openChm.UseVisualStyleBackColor = true;
		// 
		// _convertButton
		// 
		_convertButton.AutoSize = true;
		_convertButton.Location = new Point(165, 389);
		_convertButton.Name = "_convertButton";
		_convertButton.Size = new Size(94, 27);
		_convertButton.TabIndex = 18;
		_convertButton.Text = "Converti";
		_convertButton.UseVisualStyleBackColor = true;
		_convertButton.Click += Convert_Click;
		// 
		// _logLabel
		// 
		_logLabel.Anchor = AnchorStyles.Left;
		_logLabel.AutoSize = true;
		_logLabel.Location = new Point(15, 511);
		_logLabel.Name = "_logLabel";
		_logLabel.Size = new Size(30, 15);
		_logLabel.TabIndex = 19;
		_logLabel.Text = "Log:";
		_logLabel.TextAlign = ContentAlignment.MiddleLeft;
		// 
		// _logPanel
		// 
		_root.SetColumnSpan(_logPanel, 2);
		_logPanel.Controls.Add(_log);
		_logPanel.Controls.Add(_progress);
		_logPanel.Dock = DockStyle.Fill;
		_logPanel.Location = new Point(165, 422);
		_logPanel.Name = "_logPanel";
		_logPanel.Size = new Size(720, 194);
		_logPanel.TabIndex = 19;
		// 
		// _log
		// 
		_log.Dock = DockStyle.Fill;
		_log.Font = new Font("Consolas", 8.5F);
		_log.Location = new Point(0, 0);
		_log.Multiline = true;
		_log.Name = "_log";
		_log.ReadOnly = true;
		_log.ScrollBars = ScrollBars.Both;
		_log.Size = new Size(720, 171);
		_log.TabIndex = 0;
		_log.WordWrap = false;
		// 
		// _progress
		// 
		_progress.Dock = DockStyle.Bottom;
		_progress.Location = new Point(0, 171);
		_progress.Name = "_progress";
		_progress.Size = new Size(720, 23);
		_progress.Style = ProgressBarStyle.Marquee;
		_progress.TabIndex = 1;
		_progress.Visible = false;
		// 
		// _toolStrip
		// 
		_toolStrip.GripStyle = ToolStripGripStyle.Hidden;
		_toolStrip.Items.AddRange(new ToolStripItem[] { _newProjectButton, _openProjectButton, _saveProjectButton, _saveProjectAsButton });
		_toolStrip.Location = new Point(0, 0);
		_toolStrip.Name = "_toolStrip";
		_toolStrip.Padding = new Padding(8, 4, 8, 4);
		_toolStrip.RenderMode = ToolStripRenderMode.System;
		_toolStrip.Size = new Size(900, 30);
		_toolStrip.TabIndex = 0;
		// 
		// _newProjectButton
		// 
		_newProjectButton.Name = "_newProjectButton";
		_newProjectButton.Size = new Size(47, 19);
		_newProjectButton.Text = "Nuovo";
		_newProjectButton.Click += NewProject_Click;
		// 
		// _openProjectButton
		// 
		_openProjectButton.Name = "_openProjectButton";
		_openProjectButton.Size = new Size(42, 19);
		_openProjectButton.Text = "Apri...";
		_openProjectButton.Click += OpenProject_Click;
		// 
		// _saveProjectButton
		// 
		_saveProjectButton.Name = "_saveProjectButton";
		_saveProjectButton.Size = new Size(38, 19);
		_saveProjectButton.Text = "Salva";
		_saveProjectButton.Click += SaveProject_Click;
		// 
		// _saveProjectAsButton
		// 
		_saveProjectAsButton.Name = "_saveProjectAsButton";
		_saveProjectAsButton.Size = new Size(104, 19);
		_saveProjectAsButton.Text = "Salva con nome...";
		_saveProjectAsButton.Click += SaveProjectAs_Click;
		// 
		// MainForm
		// 
		AcceptButton = _convertButton;
		AutoScaleDimensions = new SizeF(7F, 15F);
		AutoScaleMode = AutoScaleMode.Font;
		ClientSize = new Size(900, 661);
		Controls.Add(_root);
		Controls.Add(_toolStrip);
		Font = new Font("Segoe UI", 9F);
		MinimumSize = new Size(760, 620);
		Name = "MainForm";
		StartPosition = FormStartPosition.CenterScreen;
		Text = "Word2Chm - Da Word a guida HTML Help (.chm)";
		_root.ResumeLayout(false);
		_root.PerformLayout();
		((System.ComponentModel.ISupportInitialize)_pageLevel).EndInit();
		((System.ComponentModel.ISupportInitialize)_bodyFontSize).EndInit();
		_optionsPanel.ResumeLayout(false);
		_optionsPanel.PerformLayout();
		_logPanel.ResumeLayout(false);
		_logPanel.PerformLayout();
		_toolStrip.ResumeLayout(false);
		_toolStrip.PerformLayout();
		ResumeLayout(false);
		PerformLayout();
	}

	#endregion

	private TableLayoutPanel _root = null!;
    private FlowLayoutPanel _optionsPanel = null!;
    private Panel _logPanel = null!;
    private ToolStrip _toolStrip = null!;
    private ToolStripButton _newProjectButton = null!;
    private ToolStripButton _openProjectButton = null!;
    private ToolStripButton _saveProjectButton = null!;
    private ToolStripButton _saveProjectAsButton = null!;
    private TextBox _docxPath = null!;
    private TextBox _outputDirectory = null!;
    private TextBox _baseName = null!;
    private TextBox _contextIdHeaderPath = null!;
    private TextBox _chmFileName = null!;
    private TextBox _chmCopyPath = null!;
    private NumericUpDown _pageLevel = null!;
    private NumericUpDown _bodyFontSize = null!;
    private TextBox _templateDirectory = null!;
    private TextBox _footer = null!;
    private TextBox _hhcPath = null!;
    private CheckBox _compileChm = null!;
    private CheckBox _openOutput = null!;
    private CheckBox _openChm = null!;
    private Button _convertButton = null!;
    private TextBox _log = null!;
    private ProgressBar _progress = null!;
    private Label _docxPathLabel = null!;
    private Button _browseDocxButton = null!;
    private Label _outputDirectoryLabel = null!;
    private Button _browseOutputButton = null!;
    private Label _baseNameLabel = null!;
    private Label _contextIdHeaderPathLabel = null!;
    private Button _browseHeaderButton = null!;
    private Label _chmFileNameLabel = null!;
    private Label _chmCopyPathLabel = null!;
    private Button _browseChmCopyButton = null!;
    private Label _pageLevelLabel = null!;
    private Label _bodyFontSizeLabel = null!;
    private Label _templateDirectoryLabel = null!;
    private Button _browseTemplateButton = null!;
    private Label _footerLabel = null!;
    private Label _hhcPathLabel = null!;
    private Button _detectHhcButton = null!;
    private Label _logLabel = null!;
}
