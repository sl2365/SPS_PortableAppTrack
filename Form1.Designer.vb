<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()

        '-- Instantiate all controls --
        Me.menuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.tsmiFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiOpenFolder = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiFileSep1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiFileSep2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiAddTrack = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiDeleteTrack = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiAbout = New System.Windows.Forms.ToolStripMenuItem()

        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.toolStripLeft = New System.Windows.Forms.ToolStrip()
        Me.tsbOpenFolder = New System.Windows.Forms.ToolStripButton()
        Me.tsSep1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbAddTrack = New System.Windows.Forms.ToolStripButton()
        Me.tsbDeleteTrack = New System.Windows.Forms.ToolStripButton()
        Me.tsSep2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbSave = New System.Windows.Forms.ToolStripButton()
        Me.lvwTracks = New System.Windows.Forms.ListView()
        Me.colName = New System.Windows.Forms.ColumnHeader()
        Me.colVersion = New System.Windows.Forms.ColumnHeader()
        Me.colStatus = New System.Windows.Forms.ColumnHeader()
        Me.statusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.sslStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.sslSpacer = New System.Windows.Forms.ToolStripStatusLabel()
        Me.sslCount = New System.Windows.Forms.ToolStripStatusLabel()

        Me.ctxListMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ctxEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctxOpenHome = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctxListSep1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ctxCopyHome = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctxCopyDownload = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctxListSep2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ctxDeleteTrack = New System.Windows.Forms.ToolStripMenuItem()

        Me.splitRight = New System.Windows.Forms.SplitContainer()
        Me.grpDetails = New System.Windows.Forms.GroupBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.txtVersion = New System.Windows.Forms.TextBox()
        Me.lblHomePage = New System.Windows.Forms.Label()
        Me.txtHomePage = New System.Windows.Forms.TextBox()
        Me.btnOpenHomePage = New System.Windows.Forms.Button()
        Me.lblDownloadLink = New System.Windows.Forms.Label()
        Me.txtDownloadLink = New System.Windows.Forms.TextBox()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()

        Me.pnlWebControls = New System.Windows.Forms.Panel()
        Me.lblUrl = New System.Windows.Forms.Label()
        Me.txtUrl = New System.Windows.Forms.TextBox()
        Me.btnDownload = New System.Windows.Forms.Button()
        Me.btnClearPage = New System.Windows.Forms.Button()
        Me.lblFind = New System.Windows.Forms.Label()
        Me.txtFind = New System.Windows.Forms.TextBox()
        Me.btnFindFirst = New System.Windows.Forms.Button()
        Me.btnFindPrev = New System.Windows.Forms.Button()
        Me.btnFindNext = New System.Windows.Forms.Button()
        Me.lblMatches = New System.Windows.Forms.Label()
        Me.rtbContent = New System.Windows.Forms.RichTextBox()

        '-- Suspend layouts --
        Me.menuStrip1.SuspendLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        Me.toolStripLeft.SuspendLayout()
        Me.statusStrip1.SuspendLayout()
        Me.ctxListMenu.SuspendLayout()
        CType(Me.splitRight, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitRight.Panel1.SuspendLayout()
        Me.splitRight.Panel2.SuspendLayout()
        Me.splitRight.SuspendLayout()
        Me.grpDetails.SuspendLayout()
        Me.pnlWebControls.SuspendLayout()
        Me.SuspendLayout()

        '======================================================
        ' menuStrip1
        '======================================================
        Me.menuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {
            Me.tsmiFile,
            Me.tsmiEdit,
            Me.tsmiHelp})
        Me.menuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.menuStrip1.Name = "menuStrip1"
        Me.menuStrip1.Size = New System.Drawing.Size(1200, 24)
        Me.menuStrip1.TabIndex = 0
        Me.menuStrip1.Text = "menuStrip1"

        ' File menu
        Me.tsmiFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {
            Me.tsmiOpenFolder,
            Me.tsmiFileSep1,
            Me.tsmiSave,
            Me.tsmiFileSep2,
            Me.tsmiExit})
        Me.tsmiFile.Name = "tsmiFile"
        Me.tsmiFile.Size = New System.Drawing.Size(37, 20)
        Me.tsmiFile.Text = "&File"

        Me.tsmiOpenFolder.Name = "tsmiOpenFolder"
        Me.tsmiOpenFolder.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.tsmiOpenFolder.Size = New System.Drawing.Size(210, 22)
        Me.tsmiOpenFolder.Text = "&Open Folder..."

        Me.tsmiFileSep1.Name = "tsmiFileSep1"
        Me.tsmiFileSep1.Size = New System.Drawing.Size(207, 6)

        Me.tsmiSave.Name = "tsmiSave"
        Me.tsmiSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.tsmiSave.Size = New System.Drawing.Size(210, 22)
        Me.tsmiSave.Text = "&Save Track"

        Me.tsmiFileSep2.Name = "tsmiFileSep2"
        Me.tsmiFileSep2.Size = New System.Drawing.Size(207, 6)

        Me.tsmiExit.Name = "tsmiExit"
        Me.tsmiExit.Size = New System.Drawing.Size(210, 22)
        Me.tsmiExit.Text = "E&xit"

        ' Edit menu
        Me.tsmiEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {
            Me.tsmiAddTrack,
            Me.tsmiDeleteTrack})
        Me.tsmiEdit.Name = "tsmiEdit"
        Me.tsmiEdit.Size = New System.Drawing.Size(39, 20)
        Me.tsmiEdit.Text = "&Edit"

        Me.tsmiAddTrack.Name = "tsmiAddTrack"
        Me.tsmiAddTrack.Size = New System.Drawing.Size(180, 22)
        Me.tsmiAddTrack.Text = "&Add Track"

        Me.tsmiDeleteTrack.Name = "tsmiDeleteTrack"
        Me.tsmiDeleteTrack.Size = New System.Drawing.Size(180, 22)
        Me.tsmiDeleteTrack.Text = "&Delete Track"

        ' Help menu
        Me.tsmiHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiAbout})
        Me.tsmiHelp.Name = "tsmiHelp"
        Me.tsmiHelp.Size = New System.Drawing.Size(44, 20)
        Me.tsmiHelp.Text = "&Help"

        Me.tsmiAbout.Name = "tsmiAbout"
        Me.tsmiAbout.Size = New System.Drawing.Size(180, 22)
        Me.tsmiAbout.Text = "&About..."

        '======================================================
        ' splitMain  (Vertical = left|right split)
        '======================================================
        Me.splitMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitMain.Location = New System.Drawing.Point(0, 24)
        Me.splitMain.Name = "splitMain"
        Me.splitMain.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.splitMain.Size = New System.Drawing.Size(1200, 726)
        Me.splitMain.SplitterDistance = 380
        Me.splitMain.SplitterWidth = 4
        Me.splitMain.TabIndex = 1

        ' splitMain.Panel1 – children (order matters: Bottom first, Top next, Fill last)
        Me.splitMain.Panel1.Controls.Add(Me.statusStrip1)
        Me.splitMain.Panel1.Controls.Add(Me.toolStripLeft)
        Me.splitMain.Panel1.Controls.Add(Me.lvwTracks)

        ' splitMain.Panel2 – children
        Me.splitMain.Panel2.Controls.Add(Me.splitRight)

        '======================================================
        ' toolStripLeft
        '======================================================
        Me.toolStripLeft.Items.AddRange(New System.Windows.Forms.ToolStripItem() {
            Me.tsbOpenFolder,
            Me.tsSep1,
            Me.tsbAddTrack,
            Me.tsbDeleteTrack,
            Me.tsSep2,
            Me.tsbSave})
        Me.toolStripLeft.Dock = System.Windows.Forms.DockStyle.Top
        Me.toolStripLeft.Location = New System.Drawing.Point(0, 0)
        Me.toolStripLeft.Name = "toolStripLeft"
        Me.toolStripLeft.Size = New System.Drawing.Size(380, 25)
        Me.toolStripLeft.TabIndex = 0
        Me.toolStripLeft.Text = "toolStripLeft"

        Me.tsbOpenFolder.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbOpenFolder.Name = "tsbOpenFolder"
        Me.tsbOpenFolder.Size = New System.Drawing.Size(76, 22)
        Me.tsbOpenFolder.Text = "Open Folder"
        Me.tsbOpenFolder.ToolTipText = "Open a folder containing SPS files (Ctrl+O)"

        Me.tsSep1.Name = "tsSep1"
        Me.tsSep1.Size = New System.Drawing.Size(6, 25)

        Me.tsbAddTrack.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbAddTrack.Name = "tsbAddTrack"
        Me.tsbAddTrack.Size = New System.Drawing.Size(32, 22)
        Me.tsbAddTrack.Text = "Add"
        Me.tsbAddTrack.ToolTipText = "Add a new track"

        Me.tsbDeleteTrack.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbDeleteTrack.Name = "tsbDeleteTrack"
        Me.tsbDeleteTrack.Size = New System.Drawing.Size(45, 22)
        Me.tsbDeleteTrack.Text = "Delete"
        Me.tsbDeleteTrack.ToolTipText = "Delete the selected track"

        Me.tsSep2.Name = "tsSep2"
        Me.tsSep2.Size = New System.Drawing.Size(6, 25)

        Me.tsbSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbSave.Name = "tsbSave"
        Me.tsbSave.Size = New System.Drawing.Size(35, 22)
        Me.tsbSave.Text = "Save"
        Me.tsbSave.ToolTipText = "Save changes to current track (Ctrl+S)"

        '======================================================
        ' lvwTracks
        '======================================================
        Me.lvwTracks.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {
            Me.colName,
            Me.colVersion,
            Me.colStatus})
        Me.lvwTracks.ContextMenuStrip = Me.ctxListMenu
        Me.lvwTracks.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvwTracks.FullRowSelect = True
        Me.lvwTracks.GridLines = True
        Me.lvwTracks.HideSelection = False
        Me.lvwTracks.Location = New System.Drawing.Point(0, 25)
        Me.lvwTracks.MultiSelect = False
        Me.lvwTracks.Name = "lvwTracks"
        Me.lvwTracks.Size = New System.Drawing.Size(380, 679)
        Me.lvwTracks.TabIndex = 1
        Me.lvwTracks.UseCompatibleStateImageBehavior = False
        Me.lvwTracks.View = System.Windows.Forms.View.Details

        Me.colName.Text = "App Name"
        Me.colName.Width = 200

        Me.colVersion.Text = "Version"
        Me.colVersion.Width = 80

        Me.colStatus.Text = "Status"
        Me.colStatus.Width = 80

        '======================================================
        ' statusStrip1
        '======================================================
        Me.statusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {
            Me.sslStatus,
            Me.sslSpacer,
            Me.sslCount})
        Me.statusStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.statusStrip1.Location = New System.Drawing.Point(0, 726)
        Me.statusStrip1.Name = "statusStrip1"
        Me.statusStrip1.Size = New System.Drawing.Size(380, 22)
        Me.statusStrip1.TabIndex = 2
        Me.statusStrip1.Text = "statusStrip1"

        Me.sslStatus.Name = "sslStatus"
        Me.sslStatus.Size = New System.Drawing.Size(200, 17)
        Me.sslStatus.Text = "Ready – Open a folder to begin"
        Me.sslStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

        Me.sslSpacer.Name = "sslSpacer"
        Me.sslSpacer.Size = New System.Drawing.Size(100, 17)
        Me.sslSpacer.Spring = True

        Me.sslCount.Name = "sslCount"
        Me.sslCount.Size = New System.Drawing.Size(55, 17)
        Me.sslCount.Text = "0 tracks"

        '======================================================
        ' ctxListMenu
        '======================================================
        Me.ctxListMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {
            Me.ctxEdit,
            Me.ctxOpenHome,
            Me.ctxListSep1,
            Me.ctxCopyHome,
            Me.ctxCopyDownload,
            Me.ctxListSep2,
            Me.ctxDeleteTrack})
        Me.ctxListMenu.Name = "ctxListMenu"
        Me.ctxListMenu.Size = New System.Drawing.Size(200, 126)

        Me.ctxEdit.Name = "ctxEdit"
        Me.ctxEdit.Size = New System.Drawing.Size(199, 22)
        Me.ctxEdit.Text = "Edit Track"

        Me.ctxOpenHome.Name = "ctxOpenHome"
        Me.ctxOpenHome.Size = New System.Drawing.Size(199, 22)
        Me.ctxOpenHome.Text = "Open Home Page in Browser"

        Me.ctxListSep1.Name = "ctxListSep1"
        Me.ctxListSep1.Size = New System.Drawing.Size(196, 6)

        Me.ctxCopyHome.Name = "ctxCopyHome"
        Me.ctxCopyHome.Size = New System.Drawing.Size(199, 22)
        Me.ctxCopyHome.Text = "Copy Home Page URL"

        Me.ctxCopyDownload.Name = "ctxCopyDownload"
        Me.ctxCopyDownload.Size = New System.Drawing.Size(199, 22)
        Me.ctxCopyDownload.Text = "Copy Download URL"

        Me.ctxListSep2.Name = "ctxListSep2"
        Me.ctxListSep2.Size = New System.Drawing.Size(196, 6)

        Me.ctxDeleteTrack.Name = "ctxDeleteTrack"
        Me.ctxDeleteTrack.Size = New System.Drawing.Size(199, 22)
        Me.ctxDeleteTrack.Text = "Delete Track"

        '======================================================
        ' splitRight  (Horizontal = top|bottom split within Panel2)
        '======================================================
        Me.splitRight.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitRight.Location = New System.Drawing.Point(0, 0)
        Me.splitRight.Name = "splitRight"
        Me.splitRight.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.splitRight.Size = New System.Drawing.Size(816, 726)
        Me.splitRight.SplitterDistance = 230
        Me.splitRight.SplitterWidth = 4
        Me.splitRight.TabIndex = 0

        ' splitRight.Panel1 – Track Details
        Me.splitRight.Panel1.Controls.Add(Me.grpDetails)

        ' splitRight.Panel2 – Web page viewer (Bottom first, then Fill)
        Me.splitRight.Panel2.Controls.Add(Me.rtbContent)
        Me.splitRight.Panel2.Controls.Add(Me.pnlWebControls)

        '======================================================
        ' grpDetails
        '======================================================
        Me.grpDetails.Controls.AddRange(New System.Windows.Forms.Control() {
            Me.lblName, Me.txtName,
            Me.lblVersion, Me.txtVersion,
            Me.lblHomePage, Me.txtHomePage, Me.btnOpenHomePage,
            Me.lblDownloadLink, Me.txtDownloadLink,
            Me.lblDescription, Me.txtDescription,
            Me.btnSave, Me.btnClear})
        Me.grpDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpDetails.Location = New System.Drawing.Point(0, 0)
        Me.grpDetails.Name = "grpDetails"
        Me.grpDetails.Size = New System.Drawing.Size(816, 230)
        Me.grpDetails.TabIndex = 0
        Me.grpDetails.TabStop = False
        Me.grpDetails.Text = "Track Details"

        ' Row 1 – App Name
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(8, 24)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(61, 13)
        Me.lblName.Text = "App Name:"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        Me.txtName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtName.Location = New System.Drawing.Point(110, 21)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(694, 20)
        Me.txtName.TabIndex = 0

        ' Row 2 – Version
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(8, 50)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(87, 13)
        Me.lblVersion.Text = "Current Version:"
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        Me.txtVersion.Location = New System.Drawing.Point(110, 47)
        Me.txtVersion.Name = "txtVersion"
        Me.txtVersion.Size = New System.Drawing.Size(130, 20)
        Me.txtVersion.TabIndex = 1

        ' Row 3 – Home Page URL
        Me.lblHomePage.AutoSize = True
        Me.lblHomePage.Location = New System.Drawing.Point(8, 76)
        Me.lblHomePage.Name = "lblHomePage"
        Me.lblHomePage.Size = New System.Drawing.Size(94, 13)
        Me.lblHomePage.Text = "Home Page URL:"
        Me.lblHomePage.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        Me.txtHomePage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHomePage.Location = New System.Drawing.Point(110, 73)
        Me.txtHomePage.Name = "txtHomePage"
        Me.txtHomePage.Size = New System.Drawing.Size(590, 20)
        Me.txtHomePage.TabIndex = 2

        Me.btnOpenHomePage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOpenHomePage.Location = New System.Drawing.Point(706, 71)
        Me.btnOpenHomePage.Name = "btnOpenHomePage"
        Me.btnOpenHomePage.Size = New System.Drawing.Size(100, 24)
        Me.btnOpenHomePage.TabIndex = 3
        Me.btnOpenHomePage.Text = "Open in Browser"
        Me.btnOpenHomePage.UseVisualStyleBackColor = True

        ' Row 4 – Download URL
        Me.lblDownloadLink.AutoSize = True
        Me.lblDownloadLink.Location = New System.Drawing.Point(8, 101)
        Me.lblDownloadLink.Name = "lblDownloadLink"
        Me.lblDownloadLink.Size = New System.Drawing.Size(88, 13)
        Me.lblDownloadLink.Text = "Download URL:"
        Me.lblDownloadLink.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        Me.txtDownloadLink.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDownloadLink.Location = New System.Drawing.Point(110, 98)
        Me.txtDownloadLink.Name = "txtDownloadLink"
        Me.txtDownloadLink.Size = New System.Drawing.Size(698, 20)
        Me.txtDownloadLink.TabIndex = 4

        ' Row 5 – Description
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Location = New System.Drawing.Point(8, 127)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(63, 13)
        Me.lblDescription.Text = "Description:"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        Me.txtDescription.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(110, 124)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(698, 20)
        Me.txtDescription.TabIndex = 5

        ' Row 6 – Buttons
        Me.btnSave.Location = New System.Drawing.Point(110, 152)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(120, 28)
        Me.btnSave.TabIndex = 6
        Me.btnSave.Text = "Save Changes"
        Me.btnSave.BackColor = System.Drawing.Color.DarkGreen
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.UseVisualStyleBackColor = False

        Me.btnClear.Location = New System.Drawing.Point(238, 152)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(100, 28)
        Me.btnClear.TabIndex = 7
        Me.btnClear.Text = "Clear / New"
        Me.btnClear.UseVisualStyleBackColor = True

        '======================================================
        ' pnlWebControls  (top of the web viewer area)
        '======================================================
        Me.pnlWebControls.Controls.AddRange(New System.Windows.Forms.Control() {
            Me.lblUrl, Me.txtUrl, Me.btnDownload, Me.btnClearPage,
            Me.lblFind, Me.txtFind, Me.btnFindFirst, Me.btnFindPrev, Me.btnFindNext, Me.lblMatches})
        Me.pnlWebControls.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlWebControls.BackColor = System.Drawing.SystemColors.ControlLight
        Me.pnlWebControls.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlWebControls.Location = New System.Drawing.Point(0, 0)
        Me.pnlWebControls.Name = "pnlWebControls"
        Me.pnlWebControls.Size = New System.Drawing.Size(816, 58)
        Me.pnlWebControls.TabIndex = 0

        ' Row 1 – URL + Download
        Me.lblUrl.AutoSize = True
        Me.lblUrl.Location = New System.Drawing.Point(5, 9)
        Me.lblUrl.Name = "lblUrl"
        Me.lblUrl.Size = New System.Drawing.Size(55, 13)
        Me.lblUrl.Text = "Page URL:"

        Me.txtUrl.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUrl.Location = New System.Drawing.Point(65, 6)
        Me.txtUrl.Name = "txtUrl"
        Me.txtUrl.Size = New System.Drawing.Size(565, 20)
        Me.txtUrl.TabIndex = 0

        Me.btnDownload.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDownload.Location = New System.Drawing.Point(636, 4)
        Me.btnDownload.Name = "btnDownload"
        Me.btnDownload.Size = New System.Drawing.Size(110, 24)
        Me.btnDownload.TabIndex = 1
        Me.btnDownload.Text = "Download Page"
        Me.btnDownload.UseVisualStyleBackColor = True

        Me.btnClearPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearPage.Location = New System.Drawing.Point(752, 4)
        Me.btnClearPage.Name = "btnClearPage"
        Me.btnClearPage.Size = New System.Drawing.Size(56, 24)
        Me.btnClearPage.TabIndex = 2
        Me.btnClearPage.Text = "Clear"
        Me.btnClearPage.UseVisualStyleBackColor = True

        ' Row 2 – Find
        Me.lblFind.AutoSize = True
        Me.lblFind.Location = New System.Drawing.Point(5, 33)
        Me.lblFind.Name = "lblFind"
        Me.lblFind.Size = New System.Drawing.Size(30, 13)
        Me.lblFind.Text = "Find:"

        Me.txtFind.Location = New System.Drawing.Point(65, 30)
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(200, 20)
        Me.txtFind.TabIndex = 3

        Me.btnFindFirst.Location = New System.Drawing.Point(271, 28)
        Me.btnFindFirst.Name = "btnFindFirst"
        Me.btnFindFirst.Size = New System.Drawing.Size(60, 24)
        Me.btnFindFirst.TabIndex = 4
        Me.btnFindFirst.Text = "|◄ First"
        Me.btnFindFirst.UseVisualStyleBackColor = True

        Me.btnFindPrev.Location = New System.Drawing.Point(336, 28)
        Me.btnFindPrev.Name = "btnFindPrev"
        Me.btnFindPrev.Size = New System.Drawing.Size(60, 24)
        Me.btnFindPrev.TabIndex = 5
        Me.btnFindPrev.Text = "◄ Prev"
        Me.btnFindPrev.UseVisualStyleBackColor = True

        Me.btnFindNext.Location = New System.Drawing.Point(401, 28)
        Me.btnFindNext.Name = "btnFindNext"
        Me.btnFindNext.Size = New System.Drawing.Size(60, 24)
        Me.btnFindNext.TabIndex = 6
        Me.btnFindNext.Text = "Next ►"
        Me.btnFindNext.UseVisualStyleBackColor = True

        Me.lblMatches.AutoSize = True
        Me.lblMatches.ForeColor = System.Drawing.Color.Gray
        Me.lblMatches.Location = New System.Drawing.Point(467, 33)
        Me.lblMatches.Name = "lblMatches"
        Me.lblMatches.Size = New System.Drawing.Size(60, 13)
        Me.lblMatches.Text = "No matches"

        '======================================================
        ' rtbContent
        '======================================================
        Me.rtbContent.BackColor = System.Drawing.Color.Black
        Me.rtbContent.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbContent.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbContent.ForeColor = System.Drawing.Color.Lime
        Me.rtbContent.Location = New System.Drawing.Point(0, 58)
        Me.rtbContent.Name = "rtbContent"
        Me.rtbContent.ReadOnly = True
        Me.rtbContent.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Both
        Me.rtbContent.Size = New System.Drawing.Size(816, 434)
        Me.rtbContent.TabIndex = 1
        Me.rtbContent.Text = ""
        Me.rtbContent.WordWrap = False

        '======================================================
        ' Form1
        '======================================================
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1200, 750)
        Me.Controls.Add(Me.splitMain)
        Me.Controls.Add(Me.menuStrip1)
        Me.MainMenuStrip = Me.menuStrip1
        Me.MinimumSize = New System.Drawing.Size(900, 500)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
        Me.Text = "SPS PortableAppTrack"

        '-- Resume layouts --
        Me.menuStrip1.ResumeLayout(False)
        Me.menuStrip1.PerformLayout()
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel1.PerformLayout()
        Me.splitMain.Panel2.ResumeLayout(False)
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.toolStripLeft.ResumeLayout(False)
        Me.toolStripLeft.PerformLayout()
        Me.statusStrip1.ResumeLayout(False)
        Me.statusStrip1.PerformLayout()
        Me.ctxListMenu.ResumeLayout(False)
        Me.splitRight.Panel1.ResumeLayout(False)
        Me.splitRight.Panel2.ResumeLayout(False)
        CType(Me.splitRight, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitRight.ResumeLayout(False)
        Me.grpDetails.ResumeLayout(False)
        Me.grpDetails.PerformLayout()
        Me.pnlWebControls.ResumeLayout(False)
        Me.pnlWebControls.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    '-- Control declarations --
    Friend WithEvents menuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents tsmiFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiOpenFolder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiFileSep1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiFileSep2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiAddTrack As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiDeleteTrack As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiAbout As System.Windows.Forms.ToolStripMenuItem

    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents toolStripLeft As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbOpenFolder As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsSep1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbAddTrack As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbDeleteTrack As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsSep2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsbSave As System.Windows.Forms.ToolStripButton
    Friend WithEvents lvwTracks As System.Windows.Forms.ListView
    Friend WithEvents colName As System.Windows.Forms.ColumnHeader
    Friend WithEvents colVersion As System.Windows.Forms.ColumnHeader
    Friend WithEvents colStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents statusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents sslStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents sslSpacer As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents sslCount As System.Windows.Forms.ToolStripStatusLabel

    Friend WithEvents ctxListMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ctxEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctxOpenHome As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctxListSep1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ctxCopyHome As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctxCopyDownload As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctxListSep2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ctxDeleteTrack As System.Windows.Forms.ToolStripMenuItem

    Friend WithEvents splitRight As System.Windows.Forms.SplitContainer
    Friend WithEvents grpDetails As System.Windows.Forms.GroupBox
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents txtVersion As System.Windows.Forms.TextBox
    Friend WithEvents lblHomePage As System.Windows.Forms.Label
    Friend WithEvents txtHomePage As System.Windows.Forms.TextBox
    Friend WithEvents btnOpenHomePage As System.Windows.Forms.Button
    Friend WithEvents lblDownloadLink As System.Windows.Forms.Label
    Friend WithEvents txtDownloadLink As System.Windows.Forms.TextBox
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button

    Friend WithEvents pnlWebControls As System.Windows.Forms.Panel
    Friend WithEvents lblUrl As System.Windows.Forms.Label
    Friend WithEvents txtUrl As System.Windows.Forms.TextBox
    Friend WithEvents btnDownload As System.Windows.Forms.Button
    Friend WithEvents btnClearPage As System.Windows.Forms.Button
    Friend WithEvents lblFind As System.Windows.Forms.Label
    Friend WithEvents txtFind As System.Windows.Forms.TextBox
    Friend WithEvents btnFindFirst As System.Windows.Forms.Button
    Friend WithEvents btnFindPrev As System.Windows.Forms.Button
    Friend WithEvents btnFindNext As System.Windows.Forms.Button
    Friend WithEvents lblMatches As System.Windows.Forms.Label
    Friend WithEvents rtbContent As System.Windows.Forms.RichTextBox

End Class
