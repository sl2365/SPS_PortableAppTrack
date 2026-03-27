<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Toggle_RightPane = New System.Windows.Forms.Button()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Label_CurrentFile = New System.Windows.Forms.Label()
        Me.Label_FindCount = New System.Windows.Forms.Label()
        Me.Label_SearchBar = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SPS_P_ListFile_Name = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SPS_P_ListView = New DoubleBufferedListView()
        Me.Track_Checked = New System.Windows.Forms.Button()
        Me.Open_File = New System.Windows.Forms.Button()
        Me.File_Save = New System.Windows.Forms.Button()
        Me.File_SaveAs = New System.Windows.Forms.Button()
        Me.Open_Browser_Border = New System.Windows.Forms.Label()
        Me.Open_Browser = New System.Windows.Forms.CheckBox()
        Me.SPS_P_Publisher_Number = New System.Windows.Forms.Label()
        Me.ReBuild_SPS_List = New System.Windows.Forms.Button()
        Me.SelectAll_CheckBox = New System.Windows.Forms.CheckBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.SPS_P_Tracker_Name = New System.Windows.Forms.TextBox()
        Me.Help = New System.Windows.Forms.Button()
        Me.SPS_P_ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolTip1Form1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStrip_OpenInBrowser = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStrip_CheckSelected = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip_UncheckSelected = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStrip_CheckAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip_UncheckAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStrip_OpenInSPSBuilder = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStrip_RefreshSelectedHash = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStrip_DeleteSelectedTrack = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LoadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveAndExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConfigToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.RestoreDefaultToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DarkModeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DarkModeToolStripMenuItem.CheckOnClick = True
        Me.DarkModeToolStripMenuItem.Name = "DarkModeToolStripMenuItem"
        Me.DarkModeToolStripMenuItem.Size = New System.Drawing.Size(187, 26)
        Me.DarkModeToolStripMenuItem.Text = "Dark Mode"
        
        ' SplitContainer and Right Pane Controls
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel_Left = New System.Windows.Forms.Panel()
        Me.Panel_Right = New System.Windows.Forms.Panel()
        
        Me.Track_URL = New System.Windows.Forms.TextBox()
        Me.Start_String = New System.Windows.Forms.TextBox()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Stop_String = New System.Windows.Forms.TextBox()
        Me.Find_String = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Download_Track_URL = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Go_To_Start_String = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Go_To_Stop_String = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Search_From_Top = New System.Windows.Forms.Button()
        Me.Save_Track = New System.Windows.Forms.Button()
        Me.Browser_TrackURL = New System.Windows.Forms.Button()
        Me.ToolTip1Form2 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Search_From_CARET = New System.Windows.Forms.Button()
        Me.Reverse_From_BOTTOM = New System.Windows.Forms.Button()
        Me.Reverse_From_CARET = New System.Windows.Forms.Button()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel_Left.SuspendLayout()
        Me.Panel_Right.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 90)
        Me.SplitContainer1.Size = New System.Drawing.Size(1397, 691)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.SplitterDistance = 450
        Me.SplitContainer1.TabIndex = 100
                
        '
        'Panel_Left (ListView pane)
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SelectAll_CheckBox)
        Me.SplitContainer1.Panel1.Controls.Add(Me.SPS_P_ListView)
        
        '
        'Panel_Right (Edit pane)
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label_FindCount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Reverse_From_CARET)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Reverse_From_BOTTOM)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Search_From_CARET)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Browser_TrackURL)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Save_Track)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Search_From_Top)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Go_To_Stop_String)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Go_To_Start_String)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Download_Track_URL)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Find_String)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Stop_String)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RichTextBox1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Start_String)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Track_URL)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        
        '
        'Label_CurrentFile
        '
        Me.Label_CurrentFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label_CurrentFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label_CurrentFile.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_CurrentFile.Location = New System.Drawing.Point(12, 30)
        Me.Label_CurrentFile.Name = "Label_CurrentFile"
        Me.Label_CurrentFile.Size = New System.Drawing.Size(1395, 55)
        Me.Label_CurrentFile.TabIndex = 101
        Me.Label_CurrentFile.Text = "Current File:"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(17, 46)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(248, 33)
        Me.ProgressBar1.TabIndex = 102
        '
        'SPS_P_ListFile_Name
        '
        Me.SPS_P_ListFile_Name.AutoSize = False
        Me.SPS_P_ListFile_Name.BackColor = System.Drawing.SystemColors.Window
        Me.SPS_P_ListFile_Name.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SPS_P_ListFile_Name.Font = New System.Drawing.Font("Arial", 10!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SPS_P_ListFile_Name.Location = New System.Drawing.Point(24, 52)
        Me.SPS_P_ListFile_Name.Name = "SPS_P_ListFile_Name"
        Me.SPS_P_ListFile_Name.Size = New System.Drawing.Size(235, 21)
        Me.SPS_P_ListFile_Name.TabIndex = 0
        Me.SPS_P_ListFile_Name.Text = "SPS_P_ListFile_Name"
        Me.ToolTip1Form2.SetToolTip(Me.SPS_P_ListFile_Name, "Current File Name" & vbCrLf & "(default: SPS_Published_Track_File.txt)")
        '
        'Open_File
        '
        Me.Open_File.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Open_File.Location = New System.Drawing.Point(272, 38)
        Me.Open_File.Name = "Open_File"
        Me.Open_File.Size = New System.Drawing.Size(40, 40)
        Me.Open_File.TabIndex = 1
        Me.Open_File.Text = ""
        Me.Open_File.UseVisualStyleBackColor = True
        Me.ToolTip1Form2.SetToolTip(Me.Open_File, "Open PAT file")
        '
        'Track_Checked
        '
        Me.Track_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Track_Checked.Location = New System.Drawing.Point(317, 38)
        Me.Track_Checked.Name = "Track_Checked"
        Me.Track_Checked.Size = New System.Drawing.Size(40, 40)
        Me.Track_Checked.TabIndex = 2
        Me.Track_Checked.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Track_Checked, "Check for updates")
        Me.Track_Checked.UseVisualStyleBackColor = True
        '
        'File_Save
        '
        Me.File_Save.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.File_Save.Location = New System.Drawing.Point(362, 38)
        Me.File_Save.Name = "File_Save"
        Me.File_Save.Size = New System.Drawing.Size(40, 40)
        Me.File_Save.TabIndex = 3
        Me.File_Save.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.File_Save, "Save changes to current PAT file")
        Me.File_Save.UseVisualStyleBackColor = True
        '
        'File_SaveAs
        '
        Me.File_SaveAs.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.File_SaveAs.Location = New System.Drawing.Point(408, 38)
        Me.File_SaveAs.Name = "File_SaveAs"
        Me.File_SaveAs.Size = New System.Drawing.Size(40, 40)
        Me.File_SaveAs.TabIndex = 4
        Me.File_SaveAs.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.File_SaveAs, "Save As...")
        Me.File_SaveAs.UseVisualStyleBackColor = True
        '
        'Open_Browser_Border
        '
        Me.Open_Browser_Border.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Open_Browser_Border.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Open_Browser_Border.Location = New System.Drawing.Point(453, 38)
        Me.Open_Browser_Border.Name = "Open_Browser_Border"
        Me.Open_Browser_Border.Size = New System.Drawing.Size(160, 39)
        Me.Open_Browser_Border.TabIndex = 105
        '
        'Open_Browser
        '
        Me.Open_Browser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Open_Browser.AutoSize = True
        Me.Open_Browser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Open_Browser.Location = New System.Drawing.Point(463, 47)
        Me.Open_Browser.Name = "Open_Browser"
        Me.Open_Browser.Size = New System.Drawing.Size(150, 38)
        Me.Open_Browser.TabIndex = 5
        Me.Open_Browser.Text = "Open in Browser"
        Me.ToolTip1Form2.SetToolTip(Me.Open_Browser, "Global Setting:" & vbcrlf & "Open selected items in default browser on Hash (Site) Change")
        Me.Open_Browser.UseVisualStyleBackColor = True

        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'SPS_P_ListView
        '
        Me.SPS_P_ListView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SPS_P_ListView.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SPS_P_ListView.GridLines = True
        Me.SPS_P_ListView.Location = New System.Drawing.Point(0, 0)
        Me.SPS_P_ListView.Name = "SPS_P_ListView"
        Me.SPS_P_ListView.Size = New System.Drawing.Size(450, 691)
        Me.SPS_P_ListView.TabIndex = 11
        Me.ToolTip1Form1.SetToolTip(Me.SPS_P_ListView, "Shift+click: Select between two points" & vbcrlf & "Ctrl+click: Select multiple items")
        Me.SPS_P_ListView.UseCompatibleStateImageBehavior = False
        Me.SPS_P_ListView.MultiSelect = True
        SPS_P_ListView.AllowColumnReorder = True
        '
        'SelectAll_CheckBox
        '
        Me.SelectAll_CheckBox.AutoSize = True
        Me.SelectAll_CheckBox.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectAll_CheckBox.Location = New System.Drawing.Point(8, 9)
        Me.SelectAll_CheckBox.Name = "SelectAll_CheckBox"
        Me.SelectAll_CheckBox.Size = New System.Drawing.Size(18, 17)
        Me.SelectAll_CheckBox.TabIndex = 10
        Me.SelectAll_CheckBox.UseVisualStyleBackColor = True
        Me.ToolTip1Form2.SetToolTip(Me.SelectAll_CheckBox, "Select all")
        
        '
        'Label_SearchBar
        '
        Me.Label_SearchBar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label_SearchBar.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_SearchBar.Location = New System.Drawing.Point(962, 32)
        Me.Label_SearchBar.Name = "Label_SearchBar"
        Me.Label_SearchBar.Size = New System.Drawing.Size(90, 18)
        Me.Label_SearchBar.TabIndex = 113
        Me.Label_SearchBar.Text = "Search Bar:"
        '
        'SPS_P_Publisher_Number
        '
        Me.SPS_P_Publisher_Number.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
'         Me.SPS_P_Publisher_Number.AutoSize = False
'         Me.SPS_P_Publisher_Number.BackColor = System.Drawing.SystemColors.Control
        Me.SPS_P_Publisher_Number.Font = New System.Drawing.Font("Arial", 7.8!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SPS_P_Publisher_Number.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.SPS_P_Publisher_Number.Location = New System.Drawing.Point(1060, 32)
        Me.SPS_P_Publisher_Number.Name = "SPS_P_Publisher_Number"
        Me.SPS_P_Publisher_Number.Size = New System.Drawing.Size(200, 18)
        Me.SPS_P_Publisher_Number.TabIndex = 106
        Me.SPS_P_Publisher_Number.Text = "SPS Files"
        Me.SPS_P_Publisher_Number.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'SPS_P_Tracker_Name
        '
        Me.SPS_P_Tracker_Name.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SPS_P_Tracker_Name.Font = New System.Drawing.Font("Arial", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SPS_P_Tracker_Name.Location = New System.Drawing.Point(970, 50)
        Me.SPS_P_Tracker_Name.Name = "SPS_P_Tracker_Name"
        Me.SPS_P_Tracker_Name.Size = New System.Drawing.Size(286, 22)
        Me.SPS_P_Tracker_Name.TabIndex = 6
        Me.SPS_P_Tracker_Name.Text = "SPS_P_Tracker_Name"
        Me.ToolTip1Form2.SetToolTip(Me.SPS_P_Tracker_Name, "SPS Publisher Name." & vbCrLf & "Add one or more names." & vbCrLf & "Leave empty to show All Publishers.")
        '
        'ReBuild_SPS_List
        '
        Me.ReBuild_SPS_List.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ReBuild_SPS_List.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReBuild_SPS_List.Location = New System.Drawing.Point(1267, 38)
        Me.ReBuild_SPS_List.Name = "ReBuild_SPS_List"
        Me.ReBuild_SPS_List.Size = New System.Drawing.Size(40, 40)
        Me.ReBuild_SPS_List.TabIndex = 7
        Me.ReBuild_SPS_List.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.ReBuild_SPS_List, "Search SPS files for the Publisher Name(s)." & vbCrLf & "If empty, gets all SPS files." & vbCrLf & "This will not erase the data written by the user.")
        Me.ReBuild_SPS_List.UseVisualStyleBackColor = True
        '
        'Toggle_RightPane
        '
        Me.Toggle_RightPane.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Toggle_RightPane.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Toggle_RightPane.Location = New System.Drawing.Point(1312, 38)
        Me.Toggle_RightPane.Name = "Toggle_RightPane"
        Me.Toggle_RightPane.Size = New System.Drawing.Size(40, 40)
        Me.Toggle_RightPane.TabIndex = 8
        Me.Toggle_RightPane.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Toggle_RightPane, "Toggle right panel visibility")
        Me.Toggle_RightPane.UseVisualStyleBackColor = True
        '
        'Help
        '
        Me.Help.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Help.Location = New System.Drawing.Point(1357, 38)
        Me.Help.Name = "Help"
        Me.Help.Size = New System.Drawing.Size(40, 40)
        Me.Help.TabIndex = 9
        Me.Help.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Help, "Opens help")
        Me.Help.UseVisualStyleBackColor = True
        '
        'SPS_P_ImageList
        '
        Me.SPS_P_ImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.SPS_P_ImageList.ImageSize = New System.Drawing.Size(16, 16)
        Me.SPS_P_ImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'ToolTip1Form1
        '
        Me.ToolTip1Form1.AutoPopDelay = 10000
        Me.ToolTip1Form1.InitialDelay = 500
        Me.ToolTip1Form1.ReshowDelay = 100
        Me.ToolTip1Form1.Tag = ""
        '
        'ToolTip1Form2
        '
        Me.ToolTip1Form2.AutoPopDelay = 10000
        Me.ToolTip1Form2.InitialDelay = 500
        Me.ToolTip1Form2.ReshowDelay = 100
        '
        'ContextMenuStrip1 - REMOVED "Edit Selected Tracks"
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() { _
            Me.ToolStrip_OpenInBrowser, Me.ToolStripSeparator3, _
            Me.ToolStrip_OpenInSPSBuilder, Me.ToolStripSeparator2, _
            Me.ToolStrip_RefreshSelectedHash, Me.ToolStripSeparator1, _
            Me.ToolStrip_DeleteSelectedTrack, Me.ToolStripSeparator7, _
            Me.ToolStrip_CheckSelected, Me.ToolStrip_UncheckSelected, Me.ToolStripSeparator8, _
            Me.ToolStrip_CheckAll, Me.ToolStrip_UncheckAll})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(279, 178)
        '
        'ToolStrip_OpenInBrowser
        '
        Me.ToolStrip_OpenInBrowser.Image = CType(resources.GetObject("ToolStrip_OpenInBrowser.Image"), System.Drawing.Image)
        Me.ToolStrip_OpenInBrowser.Name = "ToolStrip_OpenInBrowser"
        Me.ToolStrip_OpenInBrowser.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_OpenInBrowser.Text = "Open Track URL in browser"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(275, 6)
        '
        'ToolStrip_OpenInSPSBuilder
        '
        Me.ToolStrip_OpenInSPSBuilder.Image = CType(resources.GetObject("ToolStrip_OpenInSPSBuilder.Image"), System.Drawing.Image)
        Me.ToolStrip_OpenInSPSBuilder.Name = "ToolStrip_OpenInSPSBuilder"
        Me.ToolStrip_OpenInSPSBuilder.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_OpenInSPSBuilder.Text = "Open in SPS Builder"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(275, 6)
        '
        'ToolStrip_RefreshSelectedHash
        '
        Me.ToolStrip_RefreshSelectedHash.Image = CType(resources.GetObject("ToolStrip_RefreshSelectedHash.Image"), System.Drawing.Image)
        Me.ToolStrip_RefreshSelectedHash.Name = "ToolStrip_RefreshSelectedHash"
        Me.ToolStrip_RefreshSelectedHash.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_RefreshSelectedHash.Text = "Refresh Selected Hash"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(275, 6)
        '
        'ToolStrip_DeleteSelectedTrack
        '
        Me.ToolStrip_DeleteSelectedTrack.Image = CType(resources.GetObject("ToolStrip_DeleteSelectedTrack.Image"), System.Drawing.Image)
        Me.ToolStrip_DeleteSelectedTrack.Name = "ToolStrip_DeleteSelectedTrack"
        Me.ToolStrip_DeleteSelectedTrack.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_DeleteSelectedTrack.Text = "Delete Selected Tracks"
        '
        'ToolStrip_CheckSelected
        '
        Me.ToolStrip_CheckSelected.Name = "ToolStrip_CheckSelected"
        Me.ToolStrip_CheckSelected.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_CheckSelected.Text = "Check Selected"
        '
        'ToolStrip_UncheckSelected
        '
        Me.ToolStrip_UncheckSelected.Name = "ToolStrip_UncheckSelected"
        Me.ToolStrip_UncheckSelected.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_UncheckSelected.Text = "Uncheck Selected"
        '
        'ToolStrip_CheckAll
        '
        Me.ToolStrip_CheckAll.Name = "ToolStrip_CheckAll"
        Me.ToolStrip_CheckAll.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_CheckAll.Text = "Check All"
        '
        'ToolStrip_UncheckAll
        '
        Me.ToolStrip_UncheckAll.Name = "ToolStrip_UncheckAll"
        Me.ToolStrip_UncheckAll.Size = New System.Drawing.Size(278, 26)
        Me.ToolStrip_UncheckAll.Text = "Uncheck All"
        '
        'ViewSPS_NameToolStripMenuItem
        '
        Me.ViewSPS_NameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewSPS_NameToolStripMenuItem.CheckOnClick = True
        Me.ViewSPS_NameToolStripMenuItem.Checked = True
        Me.ViewSPS_NameToolStripMenuItem.Name = "ViewSPS_NameToolStripMenuItem"
        Me.ViewSPS_NameToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewSPS_NameToolStripMenuItem.Text = "SPS Name"
        '
        'ViewTrackURLToolStripMenuItem
        '
        Me.ViewTrackURLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewTrackURLToolStripMenuItem.CheckOnClick = True
        Me.ViewTrackURLToolStripMenuItem.Checked = True
        Me.ViewTrackURLToolStripMenuItem.Name = "ViewTrackURLToolStripMenuItem"
        Me.ViewTrackURLToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewTrackURLToolStripMenuItem.Text = "Track URL"
        '
        'ViewTrackStartStringToolStripMenuItem
        '
        Me.ViewTrackStartStringToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewTrackStartStringToolStripMenuItem.CheckOnClick = True
        Me.ViewTrackStartStringToolStripMenuItem.Checked = True
        Me.ViewTrackStartStringToolStripMenuItem.Name = "ViewTrackStartStringToolStripMenuItem"
        Me.ViewTrackStartStringToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewTrackStartStringToolStripMenuItem.Text = "Start String"
        '
        'ViewTrackStopStringToolStripMenuItem
        '
        Me.ViewTrackStopStringToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewTrackStopStringToolStripMenuItem.CheckOnClick = True
        Me.ViewTrackStopStringToolStripMenuItem.Checked = True
        Me.ViewTrackStopStringToolStripMenuItem.Name = "ViewTrackStopStringToolStripMenuItem"
        Me.ViewTrackStopStringToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewTrackStopStringToolStripMenuItem.Text = "Stop String"
        '
        'ViewTrackBlockHashToolStripMenuItem
        '
        Me.ViewTrackBlockHashToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewTrackBlockHashToolStripMenuItem.CheckOnClick = True
        Me.ViewTrackBlockHashToolStripMenuItem.Checked = True
        Me.ViewTrackBlockHashToolStripMenuItem.Name = "ViewTrackBlockHashToolStripMenuItem"
        Me.ViewTrackBlockHashToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewTrackBlockHashToolStripMenuItem.Text = "Track Block Hash"
        '
        'ViewVersionToolStripMenuItem
        '
        Me.ViewVersionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewVersionToolStripMenuItem.CheckOnClick = True
        Me.ViewVersionToolStripMenuItem.Checked = True
        Me.ViewVersionToolStripMenuItem.Name = "ViewVersionToolStripMenuItem"
        Me.ViewVersionToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewVersionToolStripMenuItem.Text = "Current Version"
        '
        'ViewLatestVersionToolStripMenuItem
        '
        Me.ViewLatestVersionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewLatestVersionToolStripMenuItem.CheckOnClick = True
        Me.ViewLatestVersionToolStripMenuItem.Checked = True
        Me.ViewLatestVersionToolStripMenuItem.Name = "ViewLatestVersionToolStripMenuItem"
        Me.ViewLatestVersionToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewLatestVersionToolStripMenuItem.Text = "Latest Version"
        '
        'ViewReleaseDateToolStripMenuItem
        '
        Me.ViewReleaseDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewReleaseDateToolStripMenuItem.CheckOnClick = True
        Me.ViewReleaseDateToolStripMenuItem.Checked = True
        Me.ViewReleaseDateToolStripMenuItem.Name = "ViewReleaseDateToolStripMenuItem"
        Me.ViewReleaseDateToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewReleaseDateToolStripMenuItem.Text = "Release Date"
        '
        'ViewDownloadUrlToolStripMenuItem
        '
        Me.ViewDownloadUrlToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewDownloadUrlToolStripMenuItem.CheckOnClick = True
        Me.ViewDownloadUrlToolStripMenuItem.Checked = True
        Me.ViewDownloadUrlToolStripMenuItem.Name = "ViewDownloadUrlToolStripMenuItem"
        Me.ViewDownloadUrlToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewDownloadUrlToolStripMenuItem.Text = "Download URL"
        '
        'ViewDownloadSizeKbToolStripMenuItem
        '
        Me.ViewDownloadSizeKbToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewDownloadSizeKbToolStripMenuItem.CheckOnClick = True
        Me.ViewDownloadSizeKbToolStripMenuItem.Checked = True
        Me.ViewDownloadSizeKbToolStripMenuItem.Name = "ViewDownloadSizeKbToolStripMenuItem"
        Me.ViewDownloadSizeKbToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewDownloadSizeKbToolStripMenuItem.Text = "Download Size (KB)"
        '
        'ViewSPSCreationDateToolStripMenuItem
        '
        Me.ViewSPSCreationDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewSPSCreationDateToolStripMenuItem.CheckOnClick = True
        Me.ViewSPSCreationDateToolStripMenuItem.Checked = True
        Me.ViewSPSCreationDateToolStripMenuItem.Name = "ViewSPSCreationDateToolStripMenuItem"
        Me.ViewSPSCreationDateToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewSPSCreationDateToolStripMenuItem.Text = "SPS Creation Date"
        '
        'ViewSPSModificationDateToolStripMenuItem
        '
        Me.ViewSPSModificationDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewSPSModificationDateToolStripMenuItem.CheckOnClick = True
        Me.ViewSPSModificationDateToolStripMenuItem.Checked = True
        Me.ViewSPSModificationDateToolStripMenuItem.Name = "ViewSPSModificationDateToolStripMenuItem"
        Me.ViewSPSModificationDateToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewSPSModificationDateToolStripMenuItem.Text = "SPS Modification Date"
        '
        'ViewSPSPublisherNameToolStripMenuItem
        '
        Me.ViewSPSPublisherNameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewSPSPublisherNameToolStripMenuItem.CheckOnClick = True
        Me.ViewSPSPublisherNameToolStripMenuItem.Checked = True
        Me.ViewSPSPublisherNameToolStripMenuItem.Name = "ViewSPSPublisherNameToolStripMenuItem"
        Me.ViewSPSPublisherNameToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewSPSPublisherNameToolStripMenuItem.Text = "SPS Publisher Name"
        '
        'ViewSuiteNameToolStripMenuItem
        '
        Me.ViewSuiteNameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewSuiteNameToolStripMenuItem.CheckOnClick = True
        Me.ViewSuiteNameToolStripMenuItem.Checked = True
        Me.ViewSuiteNameToolStripMenuItem.Name = "ViewSuiteNameToolStripMenuItem"
        Me.ViewSuiteNameToolStripMenuItem.Size = New System.Drawing.Size(225, 26)
        Me.ViewSuiteNameToolStripMenuItem.Text = "Suite Name"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewSPS_NameToolStripMenuItem, Me.ViewTrackURLToolStripMenuItem, Me.ViewTrackStartStringToolStripMenuItem, Me.ViewTrackStopStringToolStripMenuItem, Me.ViewTrackBlockHashToolStripMenuItem, Me.ViewVersionToolStripMenuItem, Me.ViewLatestVersionToolStripMenuItem, Me.ViewReleaseDateToolStripMenuItem, Me.ViewDownloadUrlToolStripMenuItem, Me.ViewDownloadSizeKbToolStripMenuItem, Me.ViewSPSCreationDateToolStripMenuItem, Me.ViewSPSModificationDateToolStripMenuItem, Me.ViewSPSPublisherNameToolStripMenuItem, Me.ViewSuiteNameToolStripMenuItem})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(44, 24)
        Me.ViewToolStripMenuItem.Text = "View"

        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ViewToolStripMenuItem, Me.ConfigToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1422, 28)
        Me.MenuStrip1.TabIndex = 107
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LoadToolStripMenuItem, Me.ToolStripSeparator5, Me.SaveToolStripMenuItem, Me.SaveAsToolStripMenuItem, Me.ToolStripSeparator4, Me.SaveAndExitToolStripMenuItem, Me.ToolStripSeparator6, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(44, 24)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'LoadToolStripMenuItem
        '
        Me.LoadToolStripMenuItem.Image = CType(resources.GetObject("LoadToolStripMenuItem.Image"), System.Drawing.Image)
        Me.LoadToolStripMenuItem.Name = "LoadToolStripMenuItem"
        Me.LoadToolStripMenuItem.Size = New System.Drawing.Size(172, 26)
        Me.LoadToolStripMenuItem.Text = "Open"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(169, 6)
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Image = CType(resources.GetObject("SaveToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(172, 26)
        Me.SaveToolStripMenuItem.Text = "Save"
        '
        'SaveAsToolStripMenuItem
        '
        Me.SaveAsToolStripMenuItem.Image = CType(resources.GetObject("SaveAsToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(172, 26)
        Me.SaveAsToolStripMenuItem.Text = "Save As..."
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(169, 6)
        '
        'SaveAndExitToolStripMenuItem
        '
        Me.SaveAndExitToolStripMenuItem.Name = "SaveAndExitToolStripMenuItem"
        Me.SaveAndExitToolStripMenuItem.Size = New System.Drawing.Size(172, 26)
        Me.SaveAndExitToolStripMenuItem.Text = "Save and Exit"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(169, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(172, 26)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ConfigToolStripMenuItem
        '
        Me.ConfigToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveToolStripMenuItem1, Me.RestoreDefaultToolStripMenuItem, Me.DarkModeToolStripMenuItem})
        Me.ConfigToolStripMenuItem.Name = "ConfigToolStripMenuItem"
        Me.ConfigToolStripMenuItem.Size = New System.Drawing.Size(65, 24)
        Me.ConfigToolStripMenuItem.Text = "Config"
        '
        'SaveToolStripMenuItem1
        '
        Me.SaveToolStripMenuItem1.Image = CType(resources.GetObject("SaveToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.SaveToolStripMenuItem1.Name = "SaveToolStripMenuItem1"
        Me.SaveToolStripMenuItem1.Size = New System.Drawing.Size(187, 26)
        Me.SaveToolStripMenuItem1.Text = "Save Config"
        '
        'RestoreDefaultToolStripMenuItem
        '
        Me.RestoreDefaultToolStripMenuItem.Image = CType(resources.GetObject("RestoreDefaultToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RestoreDefaultToolStripMenuItem.Name = "RestoreDefaultToolStripMenuItem"
        Me.RestoreDefaultToolStripMenuItem.Size = New System.Drawing.Size(187, 26)
        Me.RestoreDefaultToolStripMenuItem.Text = "Restore Default"
        '
        ' RIGHT PANE CONTROLS (from Form2)
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(5, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(935, 55)
        Me.Label1.TabIndex = 108
        Me.Label1.Text = "Track URL:"
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(5, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(935, 55)
        Me.Label2.TabIndex = 109
        Me.Label2.Text = "Start String:"
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(5, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(935, 55)
        Me.Label3.TabIndex = 110
        Me.Label3.Text = "Stop String:"
        '
        'Label4
        '
        Me.Label4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(5, 635)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(935, 55)
        Me.Label4.TabIndex = 111
        Me.Label4.Text = "Search:"
        '
        'Label_FindCount
        '
        Me.Label_FindCount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label_FindCount.Font = New System.Drawing.Font("Arial", 7.8!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_FindCount.BackColor = System.Drawing.Color.Green
        Me.Label_FindCount.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label_FindCount.Location = New System.Drawing.Point(62, 637)
        Me.Label_FindCount.Name = "Label_FindCount"
        Me.Label_FindCount.Size = New System.Drawing.Size(683, 18)
        Me.Label_FindCount.TabIndex = 112
        Me.Label_FindCount.Text = ""
        Me.Label_FindCount.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.Location = New System.Drawing.Point(5, 185)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(935, 440)
        Me.RichTextBox1.TabIndex = 20
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.Font = New System.Drawing.Font("Consolas", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Track_URL
        '
        Me.Track_URL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Track_URL.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Track_URL.Location = New System.Drawing.Point(15, 22)
        Me.Track_URL.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Track_URL.Name = "Track_URL"
        Me.Track_URL.Size = New System.Drawing.Size(775, 27)
        Me.Track_URL.TabIndex = 12
        Me.ToolTip1Form2.SetToolTip(Me.Track_URL, "Enter URL page for tracking the SPS app version")
        '
        'Browser_TrackURL
        '
        Me.Browser_TrackURL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Browser_TrackURL.Location = New System.Drawing.Point(800, 8)
        Me.Browser_TrackURL.Name = "Browser_TrackURL"
        Me.Browser_TrackURL.Size = New System.Drawing.Size(40, 40)
        Me.Browser_TrackURL.TabIndex = 13
        Me.Browser_TrackURL.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Browser_TrackURL, "Visit URL in your default browser")
        Me.Browser_TrackURL.UseVisualStyleBackColor = True
        '
        'Download_Track_URL
        '
        Me.Download_Track_URL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Download_Track_URL.Location = New System.Drawing.Point(845, 8)
        Me.Download_Track_URL.Name = "Download_Track_URL"
        Me.Download_Track_URL.Size = New System.Drawing.Size(40, 40)
        Me.Download_Track_URL.TabIndex = 14
        Me.Download_Track_URL.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Download_Track_URL, "Download currently selected Track URL content in the box below")
        Me.Download_Track_URL.UseVisualStyleBackColor = True
        '
        'Save_Track
        '
        Me.Save_Track.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Save_Track.Location = New System.Drawing.Point(890, 8)
        Me.Save_Track.Name = "Save_Track"
        Me.Save_Track.Size = New System.Drawing.Size(40, 40)
        Me.Save_Track.TabIndex = 15
        Me.Save_Track.Text = ""
        Me.Save_Track.UseVisualStyleBackColor = True
        Me.ToolTip1Form2.SetToolTip(Me.Save_Track, "Save changes")
        '
        'Start_String
        '
        Me.Start_String.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Start_String.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Start_String.Location = New System.Drawing.Point(15, 82)
        Me.Start_String.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Start_String.Name = "Start_String"
        Me.Start_String.Size = New System.Drawing.Size(865, 27)
        Me.Start_String.TabIndex = 16
        Me.ToolTip1Form2.SetToolTip(Me.Start_String, "Type the string where monitoring starts")
        '
        'Go_To_Start_String
        '
        Me.Go_To_Start_String.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Go_To_Start_String.Location = New System.Drawing.Point(890, 68)
        Me.Go_To_Start_String.Name = "Go_To_Start_String"
        Me.Go_To_Start_String.Size = New System.Drawing.Size(40, 40)
        Me.Go_To_Start_String.TabIndex = 17
        Me.Go_To_Start_String.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Go_To_Start_String, "Go to the Start String in the web page text")
        Me.Go_To_Start_String.UseVisualStyleBackColor = True
        '
        'Stop_String
        '
        Me.Stop_String.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Stop_String.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Stop_String.Location = New System.Drawing.Point(15, 143)
        Me.Stop_String.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Stop_String.Name = "Stop_String"
        Me.Stop_String.Size = New System.Drawing.Size(865, 27)
        Me.Stop_String.TabIndex = 18
        Me.ToolTip1Form2.SetToolTip(Me.Stop_String, "Type the string where monitoring stops")
        '
        'Go_To_Stop_String
        '
        Me.Go_To_Stop_String.Anchor = CType((System.Windows.Forms.AnchorStyles.Top _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Go_To_Stop_String.Location = New System.Drawing.Point(890, 128)
        Me.Go_To_Stop_String.Name = "Go_To_Stop_String"
        Me.Go_To_Stop_String.Size = New System.Drawing.Size(40, 40)
        Me.Go_To_Stop_String.TabIndex = 19
        Me.Go_To_Stop_String.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Go_To_Stop_String, "Go to the Stop String in the web page text")
        Me.Go_To_Stop_String.UseVisualStyleBackColor = True
        '
        'Find_String
        '
        Me.Find_String.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right)), System.Windows.Forms.AnchorStyles)
        Me.Find_String.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Find_String.Location = New System.Drawing.Point(15, 658)
        Me.Find_String.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Find_String.Name = "Find_String"
        Me.Find_String.Size = New System.Drawing.Size(730, 27)
        Me.Find_String.TabIndex = 21
        Me.ToolTip1Form2.SetToolTip(Me.Find_String, "Type text to Search page")
        '
        'Search_From_Top
        '
        Me.Search_From_Top.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Search_From_Top.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Search_From_Top.Location = New System.Drawing.Point(755, 643)
        Me.Search_From_Top.Name = "Search_From_Top"
        Me.Search_From_Top.Size = New System.Drawing.Size(40, 40)
        Me.Search_From_Top.TabIndex = 22
        Me.Search_From_Top.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Search_From_Top, "Search the first string ocurrence from TOP")
        Me.Search_From_Top.UseVisualStyleBackColor = True
        '
        'Search_From_CARET
        '
        Me.Search_From_CARET.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Search_From_CARET.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Search_From_CARET.Location = New System.Drawing.Point(800, 643)
        Me.Search_From_CARET.Name = "Search_From_CARET"
        Me.Search_From_CARET.Size = New System.Drawing.Size(40, 40)
        Me.Search_From_CARET.TabIndex = 23
        Me.Search_From_CARET.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Search_From_CARET, "Search string ocurrence from CARET")
        Me.Search_From_CARET.UseVisualStyleBackColor = True
        '
        'Reverse_From_CARET
        '
        Me.Reverse_From_CARET.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Reverse_From_CARET.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Reverse_From_CARET.Location = New System.Drawing.Point(845, 643)
        Me.Reverse_From_CARET.Name = "Reverse_From_CARET"
        Me.Reverse_From_CARET.Size = New System.Drawing.Size(40, 40)
        Me.Reverse_From_CARET.TabIndex = 24
        Me.Reverse_From_CARET.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Reverse_From_CARET, "Reverse search of the string occurence from CARET")
        Me.Reverse_From_CARET.UseVisualStyleBackColor = True
        '
        'Reverse_From_BOTTOM
        '
        Me.Reverse_From_BOTTOM.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Reverse_From_BOTTOM.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Reverse_From_BOTTOM.Location = New System.Drawing.Point(890, 643)
        Me.Reverse_From_BOTTOM.Name = "Reverse_From_BOTTOM"
        Me.Reverse_From_BOTTOM.Size = New System.Drawing.Size(40, 40)
        Me.Reverse_From_BOTTOM.TabIndex = 25
        Me.Reverse_From_BOTTOM.Text = ""
        Me.ToolTip1Form2.SetToolTip(Me.Reverse_From_BOTTOM, "Reverse search of the first string occurence from BOTTOM")
        Me.Reverse_From_BOTTOM.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopyToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(119, 30)
        Me.ContextMenuStrip2.Text = "Copy"
        '
        'CopyToolStripMenuItem
        '
        Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
        Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(118, 26)
        Me.CopyToolStripMenuItem.Text = "Copy"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1422, 795)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.Toggle_RightPane)
        Me.Controls.Add(Me.Help)
        Me.Controls.Add(Me.SPS_P_Tracker_Name)
        Me.Controls.Add(Me.ReBuild_SPS_List)
        Me.Controls.Add(Me.SPS_P_Publisher_Number)
        Me.Controls.Add(Me.Open_Browser)
        Me.Controls.Add(Me.Open_Browser_Border)
        Me.Controls.Add(Me.Open_File)
        Me.Controls.Add(Me.File_Save)
        Me.Controls.Add(Me.File_SaveAs)
        Me.Controls.Add(Me.SPS_P_ListFile_Name)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Track_Checked)
        Me.Controls.Add(Me.Label_SearchBar)
        Me.Controls.Add(Me.Label_CurrentFile)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "SPS Published App Track"
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel_Left.ResumeLayout(False)
        Me.Panel_Right.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
    
    Friend WithEvents Label_CurrentFile As System.Windows.Forms.Label
    Friend WithEvents Label_SearchBar As System.Windows.Forms.Label
    Friend WithEvents Label_FindCount As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents SPS_P_ListFile_Name As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SPS_P_ListView As DoubleBufferedListView
    Friend WithEvents Open_File As System.Windows.Forms.Button
    Friend WithEvents Track_Checked As System.Windows.Forms.Button
    Friend WithEvents File_Save As System.Windows.Forms.Button
    Friend WithEvents File_SaveAs As System.Windows.Forms.Button
    Friend WithEvents Open_Browser As System.Windows.Forms.CheckBox
    Friend WithEvents Open_Browser_Border As System.Windows.Forms.Label
    Friend WithEvents SPS_P_Publisher_Number As System.Windows.Forms.Label
    Friend WithEvents ReBuild_SPS_List As System.Windows.Forms.Button
    Friend WithEvents SelectAll_CheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents SPS_P_Tracker_Name As System.Windows.Forms.TextBox
    Friend WithEvents Toggle_RightPane As System.Windows.Forms.Button
    Friend WithEvents Help As System.Windows.Forms.Button
    Friend WithEvents SPS_P_ImageList As System.Windows.Forms.ImageList
    Friend WithEvents ToolTip1Form1 As ToolTip
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ToolStrip_OpenInSPSBuilder As ToolStripMenuItem
    Friend WithEvents ToolStrip_RefreshSelectedHash As ToolStripMenuItem
    Friend WithEvents ToolStrip_DeleteSelectedTrack As ToolStripMenuItem
    Friend WithEvents ToolStrip_OpenInBrowser As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStrip_CheckSelected As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStrip_UncheckSelected As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStrip_CheckAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStrip_UncheckAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewSPS_NameToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewTrackURLToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewTrackStartStringToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewTrackStopStringToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewTrackBlockHashToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewVersionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewLatestVersionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewReleaseDateToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewDownloadUrlToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewDownloadSizeKbToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewSPSCreationDateToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewSPSModificationDateToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewSPSPublisherNameToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewSuiteNameToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LoadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ConfigToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveAndExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents RestoreDefaultToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DarkModeToolStripMenuItem As ToolStripMenuItem
    
    ' SplitContainer
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel_Left As System.Windows.Forms.Panel
    Friend WithEvents Panel_Right As System.Windows.Forms.Panel
    
    ' Right Pane Controls (from Form2)
    Friend WithEvents Track_URL As System.Windows.Forms.TextBox
    Friend WithEvents Start_String As System.Windows.Forms.TextBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Stop_String As System.Windows.Forms.TextBox
    Friend WithEvents Find_String As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Download_Track_URL As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Go_To_Start_String As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Go_To_Stop_String As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Search_From_Top As System.Windows.Forms.Button
    Friend WithEvents Save_Track As System.Windows.Forms.Button
    Friend WithEvents Browser_TrackURL As System.Windows.Forms.Button
    Friend WithEvents ToolTip1Form2 As ToolTip
    Friend WithEvents Search_From_CARET As Button
    Friend WithEvents Reverse_From_BOTTOM As Button
    Friend WithEvents Reverse_From_CARET As Button
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents CopyToolStripMenuItem As ToolStripMenuItem
End Class