REM VVV_Easy_SyMenu "SPS_Published_Track"
REM v6.x.x.x x64 version by sl23
REM CHANGELOG:
REM 2026.03.12-v6.0.0.0: Major upgrade: Now compiled as x64. Updated text and buttons. Merged windows into a single split window. Added icons to buttons. Rearranged layout. Added listview selection methods. Added new columns. Added Column moving/sorting/sizing/hiding. Improved search functions to search all suites. Can now press Enter to search. Added DarkMode. Resized defaults. Changed settings extension to XML. Improved performance and GUI responsiveness. Updated Help file. PAT now works from any location, not just the default SyMenuSuite location. Open file now uses modern Explorer windows.
REM
REM 2025.01.14-V.5.2.0.1: Updated user agent strings. (https://useragents.io/explore)
        ' If any of your tracked URLs start failing because the server rejects your User-Agent as outdated or suspicious, you can go to that site, grab a current User-Agent string, and update the one in your code.
REM 2025.03.12-V.5.2.0.2: MERGED Form1 and Form2 into single window with SplitContainer
REM
REM 2018.03.01-V.5.2: Corrected little bugs of change the colors in second track or change the strings in edit track.
REM 2018.02.27-V.5.1: Download size test improvements (now works with SourceForge).
REM 2018.02.25-V.5.0: TLS 1.2 protocol supported (but now needed .NET Framework v.4.6.1). Download size test improvements.
REM 2017.02.26-V.4.0: Added menu bar and config file ConfigPAT.xml. Now Help opens forum Topic.
REM 2017.02.05-V.3.0: Using contextual menu and allow several Edit form.
REM 2017.02.03-V.2.0: Futherly only SPS App flavour (not Launcher needed) named SPS Published App Track (PAT).
REM                   Added SPS Builder call with local sps file copy (temporaly located in "SyMenuSuite\_Trash\_TmpPAT").
REM 2017.01.11-V.1.4b: Corrected the bug saving files with the pluging execute with Launcher (in the SPS app flavour)
REM 2017.01.09-V.1.4: Showed version in the window title. Added more search options. Corrected some bugs.
REM 2017.01.05-V.1.3: Added Tooltips.
REM                   Manage sps Or zip _Cache SPS Suite files. (SyMenu version superior To Version 5.07.6190 [2016.12.13])
REM                   Added SPS Publisher column (so the ancient <SPSPublisherName> becomes To <SPSTrackerName>).
REM                   Allows several SPS Publisher names in the SPS Tracker Name
REM 2016.12.18-V.1.2: Now in SPS stand alone program too as: SyMenu Published App Track (Others - Specialized Editors). Thanks Gian.
REM 2016.11.10-V.1.2: Corrected some bugs. Full automatic SyMenu plugin detection.
REM 2016.10.10-V.1.1: Added App Icon, Version and Release Date. Corrected some bugs. Know issue: Not automatic SyMenu plugin detection.
REM 2016.10.02-V.0.1: First published version.
REM 2016.09.21-V.Beta
REM
Option Explicit On
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Net
Imports System.Security.Cryptography
Imports System.Diagnostics

Public Class Form1
    ' Dark Mode Colours
    Private m_DarkModeBackground30 As Color = Color.FromArgb(30, 30, 30)
    Private m_DarkModeBackground45 As Color = Color.FromArgb(45, 45, 45)
    Private m_BorderColor As Color = Color.FromArgb(60, 60, 60)
    Private m_RichText As Color = Color.FromArgb(100, 150, 255)
    Private m_DarkModeRed As Color = Color.FromArgb(160, 70, 70)
    Private m_DarkModeGreen As Color = Color.FromArgb(70, 160, 70)
    Private m_DarkModeOrange As Color = Color.FromArgb(180, 130, 50)
    ' Light Mode Colours
    Private m_LightModeRed As Color = Color.LightCoral
    Private m_LightModeGreen As Color = Color.LightGreen
    Private m_LightModeOrange As Color = Color.FromArgb(255, 200, 100)
    
    Private bgWorker As New System.ComponentModel.BackgroundWorker
    Private b_FirstRun As Boolean = False
    Private b_CancelRebuild As Boolean = False
    Private img_TrackCheckedOriginal As Image = Nothing
    Private img_HTMLView As Image = Nothing
    Private img_RTFView As Image = Nothing
    Private i_RightPanelWidth As Integer = 350
    Private ReadOnly m_MsgBtnSize As New Size(60, 25)
    Private s_SyMenuSuitePath As String = Nothing
    Private i_SortColumn As Integer = -1
    Private bln_SortAscending As Boolean = True
    Private b_ShowRenderedHTML As Boolean = False
    Private b_HorizontalLayout As Boolean = False

    REM Manage the correspondence between SPS_P_ListView column header and Subitem Info
    Public Const c_SPS_Name = 0
    Public Const c_SPS_Name_Header = "     SPS Name"
    Public Const c_TrackURL = 1
    Public Const c_TrackURL_Header = "Track URL"
    Public Const c_TrackStartString = 2
    Public Const c_TrackStartString_Header = "Start String"
    Public Const c_TrackStopString = 3
    Public Const c_TrackStopString_Header = "Stop String"
    Public Const c_TrackBlockHash = 4
    Public Const c_TrackBlockHash_Header = "Track Block Hash"
    Public Const c_Version = 5
    Public Const c_Version_Header = "Version"
    Public Const c_LatestVersion = 6
    Public Const c_LatestVersion_Header = "Latest Version"
    Public Const c_ReleaseDate = 7
    Public Const c_ReleaseDate_Header = "Release Date"
    Public Const c_DownloadUrl = 8
    Public Const c_DownloadUrl_Header = "Dwnld URL"
    Public Const c_DownloadSizeKb = 9
    Public Const c_DownloadSizeKb_Header = "Dwnld Kb"
    Public Const c_SPSCreationDate = 10
    Public Const c_SPSCreationDate_Header = "SPS Creat.Date"
    Public Const c_SPSModificationDate = 11
    Public Const c_SPSModificationDate_Header = "SPS Modif.Date"
    Public Const c_SPSPublisherName = 12
    Public Const c_SPSPublisherName_Header = "SPS Publisher"
    Public Const c_SuiteName = 13
    Public Const c_SuiteName_Header = "Suite Name"
    'Global variables - Form1
    Public s_PAT_Path As String
    Public s_SyMenuSuite_Path As String
    Public s_ConfigPAT_FilePath As String
    Public s_ConfigPAT_text As String
    Public s_SPS_P_ListFile_Path As String

    'Global variables - Form2 (now integrated)
    Public i_TrackInEdit As Integer = -1
    Public s_SPSName As String = ""
    Public s_eCloseReason As String = "CloseReason.UserClosing"
    Public i_Form2_x As Integer
    Public i_Form2_y As Integer
    Public i_Form2_w As Integer
    Public i_Form2_h As Integer
    
    'Principal Form Subroutines
    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            REM Set Load Global variables and ConfigPAT
            s_PAT_Path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            'The SyMenu\ProgramFiles\SPSSuite\SyMenuSuite path to use with '\_Cache', '\_Trash', '\_SPS_Builder_sps\SPSBuilder.exe'
            Dim IniPos As Integer = InStr(1, s_PAT_Path, "\SyMenuSuite")
            If Not IniPos = 0 Then
                s_SyMenuSuite_Path = Mid(s_PAT_Path, 1, IniPos - 1) & "\SyMenuSuite"
            Else
                ' Not in normal location - try to find SyMenu
                Dim foundPath As String = FindSyMenuSuitePath()
                If Not String.IsNullOrEmpty(foundPath) Then
                    s_SyMenuSuite_Path = foundPath
                Else
                    ShowThemedMessageBox("Could not find the SyMenuSuite folder." & vbCrLf & vbCrLf &
                        "SPS Published App Track needs access to the SyMenuSuite folder to function.",
                        "SyMenuSuite not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Me.Close()
                    Return
                End If
            End If
            s_ConfigPAT_FilePath = s_PAT_Path & "\ConfigPAT.xml"
            'Get ConfigPAT.xml
            If Not File.Exists(s_ConfigPAT_FilePath) Then 'Create Default ConfigPAT.xml
                Dim s_Default_PAT_ListFile = "SPS_Published_Track_File_example.xml"  'First default List File
                If File.Exists(s_PAT_Path & "\SPS_Published_Track_File.xml") Then s_Default_PAT_ListFile = "SPS_Published_Track_File.xml" 'Second default List File
                s_ConfigPAT_text = "<PAT_ListFile>" & s_Default_PAT_ListFile & "</PAT_ListFile>" & vbCrLf &
                    "<Form1_x>" & "150" & "</Form1_x>" & vbCrLf &
                    "<Form1_y>" & "150" & "</Form1_y>" & vbCrLf &
                    "<Form1_w>" & "1100" & "</Form1_w>" & vbCrLf &
                    "<Form1_h>" & "700" & "</Form1_h>" & vbCrLf &
                    "<SplitterDistance>" & "750" & "</SplitterDistance>" & vbCrLf &
                    "<SPS_Name_w>" & "225" & "</SPS_Name_w>" & vbCrLf &
                    "<TrackURL_w>" & "130" & "</TrackURL_w>" & vbCrLf &
                    "<TrackStartString_w>" & "150" & "</TrackStartString_w>" & vbCrLf &
                    "<TrackStopString_w>" & "150" & "</TrackStopString_w>" & vbCrLf &
                    "<TrackBlockHash_w>" & "150" & "</TrackBlockHash_w>" & vbCrLf &
                    "<Version_w>" & "90" & "</Version_w>" & vbCrLf &
                    "<LatestVersion_w>" & "90" & "</LatestVersion_w>" & vbCrLf &
                    "<ReleaseDate_w>" & "120" & "</ReleaseDate_w>" & vbCrLf &
                    "<DownloadUrl_w>" & "130" & "</DownloadUrl_w>" & vbCrLf &
                    "<DownloadSizeKb_w>" & "85" & "</DownloadSizeKb_w>" & vbCrLf &
                    "<SPSCreationDate_w>" & "120" & "</SPSCreationDate_w>" & vbCrLf &
                    "<SPSModificationDate_w>" & "120" & "</SPSModificationDate_w>" & vbCrLf &
                    "<SPSPublisherName_w>" & "120" & "</SPSPublisherName_w>" & vbCrLf &
                    "<SuiteName_w>" & "120" & "</SuiteName_w>" & vbCrLf
                'Write the file name.
                File.WriteAllText(s_ConfigPAT_FilePath, s_ConfigPAT_text, Encoding.UTF8)
                s_SPS_P_ListFile_Path = s_PAT_Path & "\" & s_Default_PAT_ListFile
            Else
                s_ConfigPAT_text = File.ReadAllText(s_ConfigPAT_FilePath, Encoding.UTF8)
                s_SPS_P_ListFile_Path = s_PAT_Path & "\" & S_GetsNBlockFromText(s_ConfigPAT_text, "<PAT_ListFile>", "</PAT_ListFile>", 1)
            End If
            b_FirstRun = Not File.Exists(s_SPS_P_ListFile_Path)
            REM Principal Form position and redimension
            Me.StartPosition = FormStartPosition.CenterScreen
            Try
                Me.Left = S_GetsNBlockFromText(s_ConfigPAT_text, "<Form1_x>", "</Form1_x>", 1)
                Me.Top = S_GetsNBlockFromText(s_ConfigPAT_text, "<Form1_y>", "</Form1_y>", 1)
                Me.Width = S_GetsNBlockFromText(s_ConfigPAT_text, "<Form1_w>", "</Form1_w>", 1)
                Me.Height = S_GetsNBlockFromText(s_ConfigPAT_text, "<Form1_h>", "</Form1_h>", 1)
                Me.MinimumSize = New System.Drawing.Size(900, 500)
                Me.SplitContainer1.SplitterWidth = 5
            Catch ex As Exception
                ShowThemedMessageBox("Error setting form size: " & ex.Message, "Config Error")
            End Try
                    
            REM Restore SplitContainer position
            Dim i_SplitterDist As Integer = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SplitterDistance>", "</SplitterDistance>", 1))
            If i_SplitterDist > 100 Then
                SplitContainer1.SplitterDistance = i_SplitterDist
            End If
            SplitContainer1.FixedPanel = FixedPanel.Panel2
            
            REM Restore SplitOrientation
            Try
                Dim orientValue As String = S_GetsNBlockFromText(s_ConfigPAT_text, "<SplitOrientation>", "</SplitOrientation>", 1)
                If orientValue = "Horizontal" Then
                    b_HorizontalLayout = True
                End If
            Catch ex As Exception
                b_HorizontalLayout = False
            End Try
            
            REM Principal Form text load
            SPS_P_ListFile_Name.BackColor = Color.Empty
            SPS_P_ListFile_Name.Text = Path.GetFileName(s_SPS_P_ListFile_Path)
            SPS_P_ListFile_Name.Refresh()
            Charge_SPS_P_ListView()
            
            ' Load column visibility from config
            Try
                ViewSPS_NameToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPS_Name_w>", "</SPS_Name_w>", 1)) > 0, True, False)
                ViewTrackURLToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackURL_w>", "</TrackURL_w>", 1)) > 0, True, False)
                ViewTrackStartStringToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackStartString_w>", "</TrackStartString_w>", 1)) > 0, True, False)
                ViewTrackStopStringToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackStopString_w>", "</TrackStopString_w>", 1)) > 0, True, False)
                ViewTrackBlockHashToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackBlockHash_w>", "</TrackBlockHash_w>", 1)) > 0, True, False)
                ViewVersionToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<Version_w>", "</Version_w>", 1)) > 0, True, False)
                ViewLatestVersionToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<LatestVersion_w>", "</LatestVersion_w>", 1)) > 0, True, False)
                ViewReleaseDateToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<ReleaseDate_w>", "</ReleaseDate_w>", 1)) > 0, True, False)
                ViewDownloadUrlToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<DownloadUrl_w>", "</DownloadUrl_w>", 1)) > 0, True, False)
                ViewDownloadSizeKbToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<DownloadSizeKb_w>", "</DownloadSizeKb_w>", 1)) > 0, True, False)
                ViewSPSCreationDateToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSCreationDate_w>", "</SPSCreationDate_w>", 1)) > 0, True, False)
                ViewSPSModificationDateToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSModificationDate_w>", "</SPSModificationDate_w>", 1)) > 0, True, False)
                ViewSPSPublisherNameToolStripMenuItem.Checked = If(Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSPublisherName_w>", "</SPSPublisherName_w>", 1)) > 0, True, False)
            Catch ex As Exception
                ' If error, just leave all checked (default)
            End Try
            ' Restore column order
            Try
                Dim colOrderStr As String = S_GetsNBlockFromText(s_ConfigPAT_text, "<ColumnOrder>", "</ColumnOrder>", 1)
                If colOrderStr <> "" Then
                    Dim parts() As String = colOrderStr.Split(","c)
                    If parts.Length = SPS_P_ListView.Columns.Count Then
                        For i As Integer = 0 To parts.Length - 1
                            SPS_P_ListView.Columns(i).DisplayIndex = Convert.ToInt32(parts(i))
                        Next
                    End If
                End If
            Catch ex As Exception
                ' If error, leave default order
            End Try
            Me.SPS_P_ListView.ContextMenuStrip = ContextMenuStrip1
            
            ' Load Dark Mode setting from config
            Try
                Dim darkModeValue As String = S_GetsNBlockFromText(s_ConfigPAT_text, "<DarkMode>", "</DarkMode>", 1)
                If darkModeValue = "True" Then
                    DarkModeToolStripMenuItem.Checked = True
                    ApplyDarkMode()
                Else
                    DarkModeToolStripMenuItem.Checked = False
                    ApplyLightMode()
                End If
            Catch ex As Exception
                DarkModeToolStripMenuItem.Checked = False
                ApplyLightMode()
            End Try
            
            ' Apply layout orientation (after dark mode so theming is already set)
            If b_HorizontalLayout Then
                ApplyLayoutOrientation()
            End If
            
            REM Wire up ListView selection event to populate right pane
            AddHandler SPS_P_ListView.ItemSelectionChanged, AddressOf SPS_P_ListView_ItemSelectionChanged
            
            REM Wire up RichTextBox context menu
            Me.RichTextBox1.ContextMenuStrip = Me.ContextMenuStrip2
                        
            ' Setup BackgroundWorker for track checking
            bgWorker.WorkerReportsProgress = True
            bgWorker.WorkerSupportsCancellation = True
            AddHandler bgWorker.DoWork, AddressOf bgWorker_DoWork
            AddHandler bgWorker.ProgressChanged, AddressOf bgWorker_ProgressChanged
            AddHandler bgWorker.RunWorkerCompleted, AddressOf bgWorker_RunWorkerCompleted

            ' Improve network performance
            System.Net.ServicePointManager.DefaultConnectionLimit = 10
            
            'Load button icons
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.OpenFile.png"))
                Me.Open_File.Image = New Bitmap(img, New Size(20, 20))
                Me.Open_File.ImageAlign = ContentAlignment.MiddleCenter
                Me.Open_File.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading OpenFile icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.UpdateCheck.png"))
                Me.Track_Checked.Image = New Bitmap(img, New Size(22, 22))
                Me.Track_Checked.ImageAlign = ContentAlignment.MiddleCenter
                Me.Track_Checked.Text = ""
                img_TrackCheckedOriginal = Me.Track_Checked.Image
            Catch ex As Exception
                ShowThemedMessageBox("Error loading UpdateCheck icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Save.png"))
                Me.File_Save.Image = New Bitmap(img, New Size(20, 20))
                Me.File_Save.ImageAlign = ContentAlignment.MiddleCenter
                Me.File_Save.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Save icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Reload.png"))
                Me.ReBuild_SPS_List.Image = New Bitmap(img, New Size(20, 20))
                Me.ReBuild_SPS_List.ImageAlign = ContentAlignment.MiddleCenter
                Me.ReBuild_SPS_List.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Reload icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Sidebar.png"))
                Me.Toggle_RightPane.Image = New Bitmap(img, New Size(20, 20))
                Me.Toggle_RightPane.ImageAlign = ContentAlignment.MiddleCenter
                Me.Toggle_RightPane.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Sidebar icon: " & ex.Message)
            End Try
            Try
                Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Form1))
                Dim imgHTML As Image = CType(resources.GetObject("Toggle_HTMLView_HTML"), System.Drawing.Image)
                Dim imgRTF As Image = CType(resources.GetObject("Toggle_HTMLView_RTF"), System.Drawing.Image)
                img_HTMLView = New Bitmap(imgHTML, New Size(22, 22))
                img_RTFView = New Bitmap(imgRTF, New Size(22, 22))
                Me.Toggle_HTMLView.Image = img_HTMLView
                Me.Toggle_HTMLView.ImageAlign = ContentAlignment.MiddleCenter
                Me.Toggle_HTMLView.Text = ""
            Catch ex As Exception
                Me.Toggle_HTMLView.Text = "HTML"
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Help.png"))
                Me.Help.Image = New Bitmap(img, New Size(20, 20))
                Me.Help.ImageAlign = ContentAlignment.MiddleCenter
                Me.Help.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Help icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.SplitHorizontal.png"))
                Me.Toggle_SplitOrientation.Image = New Bitmap(img, New Size(20, 20))
                Me.Toggle_SplitOrientation.ImageAlign = ContentAlignment.MiddleCenter
                Me.Toggle_SplitOrientation.Text = ""
            Catch ex As Exception
                Me.Toggle_SplitOrientation.Text = "H/V"
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Browser.png"))
                Me.Browser_TrackURL.Image = New Bitmap(img, New Size(20, 20))
                Me.Browser_TrackURL.ImageAlign = ContentAlignment.MiddleCenter
                Me.Browser_TrackURL.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Browser icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Download.png"))
                Me.Download_Track_URL.Image = New Bitmap(img, New Size(20, 20))
                Me.Download_Track_URL.ImageAlign = ContentAlignment.MiddleCenter
                Me.Download_Track_URL.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Download icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Save.png"))
                Me.Save_Track.Image = New Bitmap(img, New Size(20, 20))
                Me.Save_Track.ImageAlign = ContentAlignment.MiddleCenter
                Me.Save_Track.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Save icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.SaveAs.png"))
                Me.File_SaveAs.Image = New Bitmap(img, New Size(20, 20))
                Me.File_SaveAs.ImageAlign = ContentAlignment.MiddleCenter
                Me.File_SaveAs.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading SaveAs icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Goto.png"))
                Me.Go_To_Start_String.Image = New Bitmap(img, New Size(20, 20))
                Me.Go_To_Start_String.ImageAlign = ContentAlignment.MiddleCenter
                Me.Go_To_Start_String.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Goto icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.Goto.png"))
                Me.Go_To_Stop_String.Image = New Bitmap(img, New Size(20, 20))
                Me.Go_To_Stop_String.ImageAlign = ContentAlignment.MiddleCenter
                Me.Go_To_Stop_String.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading Goto icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.SearchTop.png"))
                Me.Search_From_Top.Image = New Bitmap(img, New Size(22, 22))
                Me.Search_From_Top.ImageAlign = ContentAlignment.MiddleCenter
                Me.Search_From_Top.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading SearchBottom icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.SearchDown.png"))
                Me.Search_From_CARET.Image = New Bitmap(img, New Size(22, 22))
                Me.Search_From_CARET.ImageAlign = ContentAlignment.MiddleCenter
                Me.Search_From_CARET.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading SearchBottom icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.SearchUp.png"))
                Me.Reverse_From_CARET.Image = New Bitmap(img, New Size(22, 22))
                Me.Reverse_From_CARET.ImageAlign = ContentAlignment.MiddleCenter
                Me.Reverse_From_CARET.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading SearchTop icon: " & ex.Message)
            End Try
            Try
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim img As Image = Image.FromStream(assembly.GetManifestResourceStream("SPSPublishedAppTrack.SearchBottom.png"))
                Me.Reverse_From_BOTTOM.Image = New Bitmap(img, New Size(22, 22))
                Me.Reverse_From_BOTTOM.ImageAlign = ContentAlignment.MiddleCenter
                Me.Reverse_From_BOTTOM.Text = ""
            Catch ex As Exception
                ShowThemedMessageBox("Error loading SearchBottom icon: " & ex.Message)
            End Try
        Catch ex As Exception
            ShowThemedMessageBox("Error loading form: " & ex.Message & vbCrLf & ex.StackTrace, "Error")
        End Try
        ' First run - auto-rebuild to populate the list
        If b_FirstRun Then
            SPS_P_Tracker_Name.Text = ""  ' Empty search = load all SPS files
            ReBuild_SPS_List_Click(Nothing, Nothing)
            ' Auto-save the generated list
            SavePATListFile(s_SPS_P_ListFile_Path)
        End If
        File_Save.Enabled = False
        File_SaveAs.Enabled = False
        Toggle_HTMLView.Enabled = False
    End Sub
    
    Private Sub Charge_SPS_P_ListView()
        SPS_P_ListView.BeginUpdate()
        
        'Clear items and images, but keep columns intact
        SPS_P_ListView.Items.Clear()
        SPS_P_ImageList.Images.Clear()
        
        'Only set up columns and properties on first load
        If SPS_P_ListView.Columns.Count = 0 Then
            SPS_P_ListView.SmallImageList = SPS_P_ImageList
            SPS_P_ListView.CheckBoxes = True
            SPS_P_ListView.GridLines = True
            SPS_P_ListView.View = View.Details
            SPS_P_ListView.FullRowSelect = True
            SPS_P_ListView.AllowColumnReorder = False
            SPS_P_ListView.MultiSelect = True
            
            'Add column headers
            SPS_P_ListView.Columns.Add(c_SPS_Name_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPS_Name_w>", "</SPS_Name_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_TrackURL_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackURL_w>", "</TrackURL_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_TrackStartString_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackStartString_w>", "</TrackStartString_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_TrackStopString_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackStopString_w>", "</TrackStopString_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_TrackBlockHash_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackBlockHash_w>", "</TrackBlockHash_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_Version_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<Version_w>", "</Version_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_LatestVersion_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<LatestVersion_w>", "</LatestVersion_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_ReleaseDate_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<ReleaseDate_w>", "</ReleaseDate_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_DownloadUrl_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<DownloadUrl_w>", "</DownloadUrl_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_DownloadSizeKb_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<DownloadSizeKb_w>", "</DownloadSizeKb_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_SPSCreationDate_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSCreationDate_w>", "</SPSCreationDate_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_SPSModificationDate_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSModificationDate_w>", "</SPSModificationDate_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_SPSPublisherName_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSPublisherName_w>", "</SPSPublisherName_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.Columns.Add(c_SuiteName_Header, Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SuiteName_w>", "</SuiteName_w>", 1)), HorizontalAlignment.Left)
            SPS_P_ListView.AllowColumnReorder = True
        End If

        'Show the file content
        If File.Exists(s_SPS_P_ListFile_Path) Then
            Dim s_SPS_P_ListFile_text = File.ReadAllText(s_SPS_P_ListFile_Path, Encoding.UTF8)
            Dim i As Integer = 0
            Dim s_SPS_P_Info As String
            'Show Publisher
            SPS_P_Tracker_Name.Text = S_GetsNBlockFromText(s_SPS_P_ListFile_text, "<SPSTrackerName>", "</SPSTrackerName>", 1)
            SPS_P_Tracker_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeBackground45, Color.FromKnownColor(KnownColor.Window))
            SPS_P_Tracker_Name.ForeColor = If(DarkModeToolStripMenuItem.Checked, Color.White, Color.Black)
            'Populate Track file information
            s_SPS_P_Info = S_GetsNBlockFromText(s_SPS_P_ListFile_text, "<SPSPublishedEntry>", "</SPSPublishedEntry>", i + 1)
            While s_SPS_P_Info <> ""
                'Get the app name
                Dim s_SPS_Name As String = S_GetsNBlockFromText(s_SPS_P_Info, "<ProgramName>", "</ProgramName>", 1)
                'Get the app icon (16x16) <ProgramIconBase64> and set in the ImageList
                Dim img_SPS_Icon As System.Drawing.Image
                Dim s_SPS_P_IconBase64 As String = S_GetsNBlockFromText(s_SPS_P_Info, "<ProgramIconBase64>", "</ProgramIconBase64>", 1)
                Try
                    Dim imageBytes() As Byte = System.Convert.FromBase64String(s_SPS_P_IconBase64)
                    Using stream = New System.IO.MemoryStream(imageBytes, 0, imageBytes.Length)
                        img_SPS_Icon = System.Drawing.Image.FromStream(stream)
                    End Using
                Catch 'In error Blank image
                    img_SPS_Icon = New Bitmap(16, 16)
                End Try
                'Add Listview entry with s_SPS_Name as ImageKey
                Dim lvi_newItem As New ListViewItem
                lvi_newItem.SubItems(0).Text = s_SPS_Name
                lvi_newItem.SubItems(0).BackColor = Color.Empty
                lvi_newItem.UseItemStyleForSubItems = False
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<TrackURL>", "</TrackURL>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<TrackStartString>", "</TrackStartString>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<TrackStopString>", "</TrackStopString>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<TrackBlockHash>", "</TrackBlockHash>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<Version>", "</Version>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<LatestVersion>", "</LatestVersion>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<ReleaseDate>", "</ReleaseDate>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<DownloadUrl>", "</DownloadUrl>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<DownloadSizeKb>", "</DownloadSizeKb>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<SPSCreationDate>", "</SPSCreationDate>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<SPSModificationDate>", "</SPSModificationDate>", 1))
                lvi_newItem.SubItems.Add(S_GetsNBlockFromText(s_SPS_P_Info, "<SPSPublisherName>", "</SPSPublisherName>", 1))
                Dim suiteName As String = S_GetsNBlockFromText(s_SPS_P_Info, "<SuiteName>", "</SuiteName>", 1)
                lvi_newItem.SubItems.Add(Path.GetFileName(suiteName))
                ' Reconstruct full suite path from suite name
                Dim parentPath As String = Path.GetDirectoryName(s_SyMenuSuite_Path)
                lvi_newItem.Tag = Path.Combine(parentPath, Path.GetFileName(suiteName))
                lvi_newItem.ImageKey = s_SPS_Name
                SPS_P_ImageList.Images.Add(s_SPS_Name, img_SPS_Icon)
                SPS_P_ListView.Items.Add(lvi_newItem)
                'Search next app
                i = i + 1
                s_SPS_P_Info = S_GetsNBlockFromText(s_SPS_P_ListFile_text, "<SPSPublishedEntry>", "</SPSPublishedEntry>", i + 1)
            End While
            SPS_P_Publisher_Number.Text = SPS_P_ListView.Items.Count & " SPS Files"
        End If
        
        SPS_P_ListView.EndUpdate()
    End Sub
            
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPS_P_Tracker_Name.TextChanged
        SPS_P_Tracker_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
        File_Save.Enabled = True
        File_SaveAs.Enabled = True
    End Sub
    
    '======================================
    ' NEW EVENT: ListView Selection Changed
    '======================================
    Private Sub SPS_P_ListView_ItemSelectionChanged(ByVal sender As Object, ByVal e As ListViewItemSelectionChangedEventArgs) Handles SPS_P_ListView.ItemSelectionChanged
        If e.IsSelected Then
            Me.Toggle_HTMLView.Enabled = True
            RichTextBox1.Tag = Nothing
            i_TrackInEdit = e.ItemIndex
            s_SPSName = SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_SPS_Name).Text
            ' Show the text fields immediately, download web page in background
            Dim s_TrackURL As String = SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackURL).Text
            Dim s_StartString As String = SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackStartString).Text
            Dim s_StopString As String = SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackStopString).Text
            Me.Track_URL.BackColor = Color.Empty
            Me.Track_URL.Text = s_TrackURL
            Me.Start_String.BackColor = Color.Empty
            Me.Start_String.Text = s_StartString
            Me.Stop_String.BackColor = Color.Empty
            Me.Stop_String.Text = s_StopString
            Me.RichTextBox1.Clear()
            Me.RichTextBox1.Text = "Loading..."
            Put_WebPage_in_Form_Async(s_TrackURL, s_StartString, s_StopString)
        End If
    End Sub
        
    Private Sub SPS_P_ListView_MouseDown(sender As Object, e As MouseEventArgs) Handles SPS_P_ListView.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim item As ListViewItem = SPS_P_ListView.GetItemAt(e.X, e.Y)
            ' Always show context menu, but disable item-specific options when clicking empty space
            ToolStrip_OpenInBrowser.Enabled = (item IsNot Nothing)
            ToolStrip_OpenInSPSBuilder.Enabled = (item IsNot Nothing)
            ToolStrip_OpenCheckedSPS.Enabled = (SPS_P_ListView.CheckedItems.Count > 0)
            ToolStrip_RefreshSelectedHash.Enabled = (item IsNot Nothing)
            ToolStrip_DeleteSelectedTrack.Enabled = (item IsNot Nothing)
            ToolStrip_CheckSelected.Enabled = (SPS_P_ListView.SelectedItems.Count > 0)
            ToolStrip_UncheckSelected.Enabled = (SPS_P_ListView.SelectedItems.Count > 0)
            SPS_P_ListView.ContextMenuStrip = ContextMenuStrip1
        End If
    End Sub
    
    '======================================
    ' INTEGRATED Form2 CODE - STARTS HERE
    '======================================
    Public Sub Put_SPSTrack_in_Form(ByVal _SPS_order As Integer)
        Dim s_TrackURL As String = SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackURL).Text
        Dim s_StartString As String = SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackStartString).Text
        Dim s_StopString As String = SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackStopString).Text
        ' Write track values in form
        Me.Track_URL.BackColor = Color.Empty
        Me.Track_URL.Clear()
        Me.Track_URL.Text = s_TrackURL
        Me.Start_String.BackColor = Color.Empty
        Me.Start_String.Clear()
        Me.Start_String.Text = s_StartString
        Me.Stop_String.BackColor = Color.Empty
        Me.Stop_String.Clear()
        Me.Stop_String.Text = s_StopString
        Put_WebPage_in_Form(s_TrackURL, s_StartString, s_StopString)
    End Sub
    
    Private Sub Put_WebPage_in_Form(ByVal s_TrackURL As String, ByVal s_StartString As String, ByVal s_StopString As String)
        Me.RichTextBox1.Clear()
        'Download 
        Dim s_Web_Page As String
        Dim i_Start_String_Position As Integer
        Try
            Dim strOutput As String = ""
            Dim tempCookies As New CookieContainer
            Dim wrRequest As HttpWebRequest = CType(HttpWebRequest.Create(s_TrackURL), HttpWebRequest)
            wrRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0"
            wrRequest.Timeout = 10000
            wrRequest.CookieContainer = tempCookies
            Dim wrResponse As HttpWebResponse = CType(wrRequest.GetResponse(), HttpWebResponse)
            Using sr As New StreamReader(wrResponse.GetResponseStream())
                s_Web_Page = sr.ReadToEnd()
                'Close And clean up the StreamReader
                sr.Close()
                wrResponse.Close()
            End Using
            Me.Track_URL.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
            i_Start_String_Position = InStr(1, s_Web_Page, s_StartString)
            If i_Start_String_Position = 0 Then
                Me.Start_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Else
                Me.Start_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
            End If
            If InStr(i_Start_String_Position + 1, s_Web_Page, s_StopString) = 0 Then
                Me.Stop_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Else
                Me.Stop_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
            End If
            REM Process web with Web Download successful
            RichTextBox1.DetectUrls = False
            RichTextBox1.ReadOnly = True
            RichTextBox1.Text = s_Web_Page
            RichTextBox1.Tag = s_Web_Page
            REM Process block
            HighLigth_TrackBlock_in_RichTextBox(s_StartString, s_StopString)
        Catch ex As Exception
            REM Web process error
            Beep()
            Me.Track_URL.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Me.Start_String.BackColor = Color.Empty
            Me.Stop_String.BackColor = Color.Empty
        End Try
    End Sub
    
    Private Async Sub Put_WebPage_in_Form_Async(ByVal s_TrackURL As String, ByVal s_StartString As String, ByVal s_StopString As String)
        Dim s_Web_Page As String = ""
        Dim b_Success As Boolean = False
        
        Try
            ' Download on background thread - UI stays responsive
            s_Web_Page = Await Task.Run(Function()
                Dim tempCookies As New CookieContainer
                Dim wrRequest As HttpWebRequest = CType(HttpWebRequest.Create(s_TrackURL), HttpWebRequest)
                wrRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0"
                wrRequest.Timeout = 10000
                wrRequest.CookieContainer = tempCookies
                Dim wrResponse As HttpWebResponse = CType(wrRequest.GetResponse(), HttpWebResponse)
                Using sr As New StreamReader(wrResponse.GetResponseStream())
                    Dim result As String = sr.ReadToEnd()
                    sr.Close()
                    wrResponse.Close()
                    Return result
                End Using
            End Function)
            b_Success = True
        Catch ex As Exception
            b_Success = False
        End Try
        
        ' Back on UI thread now - safe to update controls
        If b_Success Then
            Me.Track_URL.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
            Dim i_Start_String_Position As Integer = InStr(1, s_Web_Page, s_StartString)
            If i_Start_String_Position = 0 Then
                Me.Start_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Else
                Me.Start_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
            End If
            If InStr(i_Start_String_Position + 1, s_Web_Page, s_StopString) = 0 Then
                Me.Stop_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Else
                Me.Stop_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
            End If
            RichTextBox1.DetectUrls = False
            RichTextBox1.ReadOnly = True
            RichTextBox1.Text = s_Web_Page
            RichTextBox1.Tag = s_Web_Page
            If b_ShowRenderedHTML Then
                Dim stripped As String = StripHtmlTags(s_Web_Page)
                RichTextBox1.Text = stripped
                Toggle_HTMLView.Image = img_RTFView
            End If
            HighLigth_TrackBlock_in_RichTextBox(s_StartString, s_StopString)
        Else
            Beep()
            Me.Track_URL.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Me.Start_String.BackColor = Color.Empty
            Me.Stop_String.BackColor = Color.Empty
            Me.RichTextBox1.Text = "Failed to download page."
        End If
    End Sub
    
    Private Sub HighLigth_TrackBlock_in_RichTextBox(ByVal s_StartString As String, ByVal s_StartStop As String)
        If Not RichTextBox1.Text.Length = 0 Then
            Dim IniPos As Integer = InStr(1, RichTextBox1.Text, s_StartString)
            If IniPos = 0 Then 'Set IniPos at begining
                IniPos = 1
            Else 'TrackStartString in Bold-Green
                RichTextBox1.Select(IniPos - 1, s_StartString.Length)
                RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
                RichTextBox1.SelectionColor = Color.Green
                RichTextBox1.ScrollToCaret()
                IniPos = IniPos + s_StartString.Length
            End If
            Dim EndPos As Integer = InStr(IniPos, RichTextBox1.Text, s_StartStop)
            If EndPos = 0 Or EndPos = 1 Then 'Set EndPos as final position
                EndPos = RichTextBox1.Text.Length
            Else 'TrackStartString in Bold-Red
                RichTextBox1.Select(EndPos - 1, s_StartStop.Length)
                RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
                RichTextBox1.SelectionColor = Color.Red
            End If
            'Track block with in Bold-Blue
            RichTextBox1.Select(IniPos - 1, EndPos - IniPos)
            RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
            RichTextBox1.SelectionColor = If(DarkModeToolStripMenuItem.Checked, m_RichText, Color.Blue)
        End If
    End Sub
    
    Private Sub Save_Track_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save_Track.Click
        'To avoid reorder problemas we must search the Track by name.
        Dim AddSPS As Boolean = True
        For i_TrackInEdit = 0 To SPS_P_ListView.Items.Count - 1
            If SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_SPS_Name).Text = s_SPSName Then
                AddSPS = False
                SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackURL).BackColor = Color.Empty
                SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackURL).Text = Track_URL.Text
                SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackStartString).Text = Start_String.Text
                SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackStopString).Text = Stop_String.Text
                SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackBlockHash).BackColor = Color.Empty
                SPS_P_ListView.Items(i_TrackInEdit).SubItems(c_TrackBlockHash).Text = ""
                SPS_P_ListFile_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
                SPS_P_ListFile_Name.Refresh()
                File_Save.Enabled = True
                File_SaveAs.Enabled = True
                Put_SPSTrack_in_Form(i_TrackInEdit)
                Exit For
            End If
        Next
        If AddSPS = True Then 'New SPS
            ShowThemedMessageBox("SPS App name not found." & vbCrLf &
                                "Rebuild for save the Track", "SPS App name", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    
    Private Sub File_SaveAs_Click(sender As Object, e As EventArgs) Handles File_SaveAs.Click
        Me.SaveFileDialog1.FileName = Path.GetFileNameWithoutExtension(s_SPS_P_ListFile_Path)
        Me.SaveFileDialog1.InitialDirectory = Path.GetDirectoryName(s_SPS_P_ListFile_Path)
        Me.SaveFileDialog1.DefaultExt = ".xml"
        Me.SaveFileDialog1.AddExtension = True
        Me.SaveFileDialog1.Filter = "Text Files|*.xml"
        Me.SaveFileDialog1.CheckFileExists = False
        Me.SaveFileDialog1.AutoUpgradeEnabled = False
        Dim result As DialogResult = Me.SaveFileDialog1.ShowDialog()
        If result = DialogResult.OK Then
            s_SPS_P_ListFile_Path = SaveFileDialog1.FileName
            SavePATListFile(s_SPS_P_ListFile_Path)
        End If
    End Sub
    
    Private Sub Download_Track_URL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Download_Track_URL.Click
        Put_WebPage_in_Form(Me.Track_URL.Text, Me.Start_String.Text, Me.Stop_String.Text)
    End Sub
    
    Private Sub Go_To_Start_String_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go_To_Start_String.Click
        If Not RichTextBox1.Text.Length = 0 Then
            Dim IniPos As Integer = InStr(1, RichTextBox1.Text, Start_String.Text)
            If IniPos = 0 Then 'Set IniPos at begining
                ShowThemedMessageBox("Start String not found!", "Go To Start String")
                Me.Start_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Else 'Redraw with the present form values
                Put_WebPage_in_Form(Me.Track_URL.Text, Me.Start_String.Text, Me.Stop_String.Text)
            End If
        End If
    End Sub
    
    Private Sub Go_To_Stop_String_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go_To_Stop_String.Click
        If Not RichTextBox1.Text.Length = 0 Then
            Dim IniPos As Integer = InStr(1, RichTextBox1.Text, Start_String.Text)
            If IniPos = 0 Then IniPos = 1 'Set IniPos at begining
            IniPos = IniPos + Start_String.Text.Length
            Dim EndPos As Integer = InStr(IniPos, RichTextBox1.Text, Stop_String.Text)
            If EndPos = 0 Then 'Set IniPos at begining
                ShowThemedMessageBox("Stop String not found after Start Stop!", "Go To Stop String")
                Me.Stop_String.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            Else 'Redraw with the present form values
                Put_WebPage_in_Form(Me.Track_URL.Text, Me.Start_String.Text, Me.Stop_String.Text)
                RichTextBox1.Select(EndPos - 1, Stop_String.Text.Length)
                RichTextBox1.ScrollToCaret()
            End If
        End If
    End Sub
    
    Private Sub Browser_TrackURL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Browser_TrackURL.Click
        Using BrowserProcess As New Process()
            With BrowserProcess.StartInfo
                .FileName = Me.Track_URL.Text
                .CreateNoWindow = True
                .UseShellExecute = True
                .WindowStyle = ProcessWindowStyle.Hidden
            End With
            BrowserProcess.Start()
        End Using
    End Sub
    
    Private Sub Search_From_Top_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Search_From_Top.Click
        If Not RichTextBox1.Text.Length = 0 Then
            UpdateFindCount()
            Dim IniPos As Integer = InStr(1, RichTextBox1.Text, Find_String.Text)
            If IniPos = 0 Then 'String not found
                ShowThemedMessageBox("String not found!", "Find String")
            Else 'Found String in Bold-Underline-black
                RichTextBox1.Select(IniPos - 1, Find_String.Text.Length)
                RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold Or FontStyle.Underline)
                RichTextBox1.SelectionColor = If(DarkModeToolStripMenuItem.Checked, Color.Yellow, Color.Orange)
                RichTextBox1.ScrollToCaret()
            End If
        End If
    End Sub
    
    Private Sub Search_From_CARET_Click(sender As Object, e As EventArgs) Handles Search_From_CARET.Click
        If Not RichTextBox1.Text.Length = 0 Then
            UpdateFindCount()
            Dim IniPos As Integer = InStr(RichTextBox1.SelectionStart + RichTextBox1.SelectionLength, RichTextBox1.Text, Find_String.Text)
            If IniPos = 0 Then 'String not found
                ShowThemedMessageBox("String not found!", "Find String")
            Else 'Found String in Bold-Underline-black
                RichTextBox1.Select(IniPos - 1, Find_String.Text.Length)
                RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold Or FontStyle.Underline)
                RichTextBox1.SelectionColor = If(DarkModeToolStripMenuItem.Checked, Color.Yellow, Color.Orange)
                RichTextBox1.ScrollToCaret()
            End If
        End If
    End Sub
    
    Private Sub Reverse_From_CARET_Click(sender As Object, e As EventArgs) Handles Reverse_From_CARET.Click
        If Not RichTextBox1.Text.Length = 0 Then
            UpdateFindCount()
            Dim IniPos As Integer = InStrRev(RichTextBox1.Text, Find_String.Text, RichTextBox1.SelectionStart)
            If IniPos = 0 Then 'String not found
                ShowThemedMessageBox("String not found!", "Find String")
            Else 'Found String in Bold-Underline-black
                RichTextBox1.Select(IniPos - 1, Find_String.Text.Length)
                RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold Or FontStyle.Underline)
                RichTextBox1.SelectionColor = If(DarkModeToolStripMenuItem.Checked, Color.Yellow, Color.Orange)
                RichTextBox1.ScrollToCaret()
            End If
        End If
    End Sub
    
    Private Sub Reverse_From_BOTTOM_Click(sender As Object, e As EventArgs) Handles Reverse_From_BOTTOM.Click
        If Not RichTextBox1.Text.Length = 0 Then
            UpdateFindCount()
            Dim IniPos As Integer = InStrRev(RichTextBox1.Text, Find_String.Text, -1)
            If IniPos = 0 Then 'String not found
                ShowThemedMessageBox("String not found!", "Find String")
            Else 'Found String in Bold-Underline-black
                RichTextBox1.Select(IniPos - 1, Find_String.Text.Length)
                RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold Or FontStyle.Underline)
                RichTextBox1.SelectionColor = If(DarkModeToolStripMenuItem.Checked, Color.Yellow, Color.Orange)
                RichTextBox1.ScrollToCaret()
            End If
        End If
    End Sub
    
    Private Function FindSyMenuSuitePath() As String
        ' First try relative to exe location (normal install)
        Dim exeDir As String = Path.GetDirectoryName(Application.ExecutablePath)
        Dim IniPos As Integer = InStr(1, exeDir, "\SyMenuSuite")
        If Not IniPos = 0 Then
            Dim normalPath As String = Mid(exeDir, 1, IniPos - 1) & "\SyMenuSuite"
            If Directory.Exists(normalPath) Then
                Return normalPath
            End If
        End If

        ' Check if we saved a path previously in config
        Dim configPath As String = exeDir & "\ConfigPAT.xml"
        If File.Exists(configPath) Then
            Dim configText As String = File.ReadAllText(configPath)
            Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(configText, "<SyMenuSuitePath>(.*?)</SyMenuSuitePath>")
            If match.Success AndAlso Directory.Exists(match.Groups(1).Value) Then
                Return match.Groups(1).Value
            End If
        End If

        ' Let user browse
        Dim fbd As New FolderBrowserDialog()
        fbd.Description = "Could not find SyMenuSuite folder." & vbCrLf & "Please browse to the SyMenuSuite folder (e.g. X:\SyMenu\ProgramFiles\SPSSuite\SyMenuSuite)"
        fbd.ShowNewFolderButton = False
        Dim topForm As New Form()
        topForm.TopMost = True
        topForm.StartPosition = FormStartPosition.CenterScreen
        topForm.Width = 0
        topForm.Height = 0
        topForm.FormBorderStyle = FormBorderStyle.None
        topForm.ShowInTaskbar = False
        topForm.Show()
        If fbd.ShowDialog(topForm) = DialogResult.OK Then
            topForm.Close()
            topForm.Dispose()
            If Directory.Exists(fbd.SelectedPath) Then
                Return fbd.SelectedPath
            End If
        Else
            topForm.Close()
            topForm.Dispose()
        End If

        Return Nothing
    End Function

    Private Function ShowThemedMessageBox(message As String, Optional title As String = "Message", Optional buttons As MessageBoxButtons = MessageBoxButtons.OK, Optional icon As MessageBoxIcon = MessageBoxIcon.None) As DialogResult
        Dim msg As New Form()
        msg.Text = title
        msg.Width = 400
        msg.Height = 180
        msg.FormBorderStyle = FormBorderStyle.FixedDialog
        msg.StartPosition = FormStartPosition.Manual
        msg.Location = New Point(Me.Left + (Me.Width - msg.Width) \ 2, Me.Top + (Me.Height - msg.Height) \ 2)
        msg.MaximizeBox = False
        msg.MinimizeBox = False
        msg.ShowInTaskbar = False

        Dim lbl As New Label()
        lbl.Text = message
        lbl.AutoSize = False
        lbl.TextAlign = ContentAlignment.MiddleCenter
        lbl.Location = New Point(10, 10)
        lbl.Size = New Size(msg.ClientSize.Width - 20, 80)
        msg.Controls.Add(lbl)

        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.FlowDirection = FlowDirection.LeftToRight
        buttonPanel.AutoSize = True
        buttonPanel.Anchor = AnchorStyles.Bottom
        msg.Controls.Add(buttonPanel)

        Dim btnSize As Size = m_MsgBtnSize

        Select Case buttons
            Case MessageBoxButtons.OK
                Dim btnOK As New Button()
                btnOK.Text = "OK"
                btnOK.Size = btnSize
                btnOK.DialogResult = DialogResult.OK
                buttonPanel.Controls.Add(btnOK)
                msg.AcceptButton = btnOK

            Case MessageBoxButtons.YesNo
                Dim btnYes As New Button()
                btnYes.Text = "Yes"
                btnYes.Size = btnSize
                btnYes.DialogResult = DialogResult.Yes
                buttonPanel.Controls.Add(btnYes)
                msg.AcceptButton = btnYes
                Dim btnNo As New Button()
                btnNo.Text = "No"
                btnNo.Size = btnSize
                btnNo.DialogResult = DialogResult.No
                buttonPanel.Controls.Add(btnNo)

            Case MessageBoxButtons.YesNoCancel
                Dim btnYes As New Button()
                btnYes.Text = "Yes"
                btnYes.Size = btnSize
                btnYes.DialogResult = DialogResult.Yes
                buttonPanel.Controls.Add(btnYes)
                msg.AcceptButton = btnYes
                Dim btnNo As New Button()
                btnNo.Text = "No"
                btnNo.Size = btnSize
                btnNo.DialogResult = DialogResult.No
                buttonPanel.Controls.Add(btnNo)
                Dim btnCancel As New Button()
                btnCancel.Text = "Cancel"
                btnCancel.Size = btnSize
                btnCancel.DialogResult = DialogResult.Cancel
                buttonPanel.Controls.Add(btnCancel)
                msg.CancelButton = btnCancel

            Case MessageBoxButtons.OKCancel
                Dim btnOK As New Button()
                btnOK.Text = "OK"
                btnOK.Size = btnSize
                btnOK.DialogResult = DialogResult.OK
                buttonPanel.Controls.Add(btnOK)
                msg.AcceptButton = btnOK
                Dim btnCancel As New Button()
                btnCancel.Text = "Cancel"
                btnCancel.Size = btnSize
                btnCancel.DialogResult = DialogResult.Cancel
                buttonPanel.Controls.Add(btnCancel)
                msg.CancelButton = btnCancel
        End Select

        ' Center the button panel
        buttonPanel.Location = New Point((msg.ClientSize.Width - buttonPanel.PreferredSize.Width) \ 2, 100)

        ' Apply theme
        If DarkModeToolStripMenuItem.Checked Then
            msg.BackColor = m_DarkModeBackground30
            lbl.ForeColor = Color.White
            For Each ctrl As Control In buttonPanel.Controls
                If TypeOf ctrl Is Button Then
                    ctrl.BackColor = m_DarkModeBackground45
                    ctrl.ForeColor = Color.White
                    CType(ctrl, Button).FlatStyle = FlatStyle.Flat
                    CType(ctrl, Button).FlatAppearance.BorderColor = m_BorderColor
                End If
            Next
        End If

        Dim result As DialogResult = msg.ShowDialog(Me)
        msg.Dispose()
        Return result
    End Function
    
    Private Function CountOccurrences(ByVal source As String, ByVal searchTerm As String) As Integer
        If String.IsNullOrEmpty(source) OrElse String.IsNullOrEmpty(searchTerm) Then Return 0
        Dim count As Integer = 0
        Dim pos As Integer = 0
        Do
            pos = InStr(pos + 1, source, searchTerm)
            If pos > 0 Then
                count += 1
            End If
        Loop While pos > 0
        Return count
    End Function

    Private Sub UpdateFindCount()
        If RichTextBox1.Text.Length > 0 AndAlso Find_String.Text.Length > 0 Then
            Dim count As Integer = CountOccurrences(RichTextBox1.Text, Find_String.Text)
            Label_FindCount.Text = count & " match" & If(count <> 1, "es", "") & " found"
        Else
            Label_FindCount.Text = ""
        End If
    End Sub
    
    Private Sub Find_String_KeyDown(sender As Object, e As KeyEventArgs) Handles Find_String.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If Not RichTextBox1.Text.Length = 0 Then
                UpdateFindCount()
                Dim startPos As Integer = RichTextBox1.SelectionStart + RichTextBox1.SelectionLength + 1
                If startPos < 1 Then startPos = 1
                Dim IniPos As Integer = InStr(startPos, RichTextBox1.Text, Find_String.Text)
                If IniPos = 0 Then
                    ' Wrap around to top
                    IniPos = InStr(1, RichTextBox1.Text, Find_String.Text)
                    If IniPos = 0 Then
                        ShowThemedMessageBox("String not found!", "Find String")
                        Return
                    End If
                End If
                RichTextBox1.Select(IniPos - 1, Find_String.Text.Length)
                RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold Or FontStyle.Underline)
                RichTextBox1.SelectionColor = If(DarkModeToolStripMenuItem.Checked, Color.Yellow, Color.Black)
                RichTextBox1.ScrollToCaret()
            End If
        End If
    End Sub
    
    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        Clipboard.SetText(RichTextBox1.SelectedText)
    End Sub
    
    Private i_SavedSplitterDistance As Integer = 450
    Private i_MaxSplitterDistance As Integer = 250  ' Add this

    Private Sub Toggle_RightPane_Click(sender As Object, e As EventArgs) Handles Toggle_RightPane.Click
        If Me.SplitContainer1.Panel2.Visible Then
            ' Hide panel and save position
            i_SavedSplitterDistance = Me.SplitContainer1.SplitterDistance
            Me.SplitContainer1.Panel2.Visible = False
            i_MaxSplitterDistance = 0
            If b_HorizontalLayout Then
                Me.SplitContainer1.SplitterDistance = Me.SplitContainer1.Height
            Else
                Me.SplitContainer1.SplitterDistance = Me.SplitContainer1.Width
            End If
            Me.SplitContainer1.IsSplitterFixed = True  ' Disable dragging
        Else
            ' Show panel and restore position
            Me.SplitContainer1.Panel2.Visible = True
            i_MaxSplitterDistance = 250
            Me.SplitContainer1.SplitterDistance = i_SavedSplitterDistance
            Me.SplitContainer1.IsSplitterFixed = False  ' Enable dragging
        End If
    End Sub

    Private Sub SplitContainer1_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer1.SplitterMoved
        ' Don't apply limits if panel is hidden
        If Not Me.SplitContainer1.Panel2.Visible Then
            Exit Sub
        End If
        
        Dim minDistance As Integer = 200
        Dim containerSize As Integer = If(b_HorizontalLayout, Me.SplitContainer1.Height, Me.SplitContainer1.Width)
        Dim maxDistance As Integer = containerSize - i_MaxSplitterDistance
        
        If Me.SplitContainer1.SplitterDistance < minDistance Then
            Me.SplitContainer1.SplitterDistance = minDistance
        ElseIf Me.SplitContainer1.SplitterDistance > maxDistance Then
            Me.SplitContainer1.SplitterDistance = maxDistance
        End If
    End Sub

    Private Sub Toggle_HTMLView_Click(sender As Object, e As EventArgs) Handles Toggle_HTMLView.Click
        If RichTextBox1.Text.Length = 0 Then Return
        b_ShowRenderedHTML = Not b_ShowRenderedHTML
        If b_ShowRenderedHTML Then
            Toggle_HTMLView.Image = img_RTFView
            ' Store raw HTML in Tag, then show stripped version
            If RichTextBox1.Tag Is Nothing OrElse RichTextBox1.Tag.ToString() = "" Then
                RichTextBox1.Tag = RichTextBox1.Text
            End If
            Dim stripped As String = StripHtmlTags(RichTextBox1.Tag.ToString())
            RichTextBox1.Text = stripped
            HighLigth_TrackBlock_in_RichTextBox(Start_String.Text, Stop_String.Text)
        Else
            Toggle_HTMLView.Image = img_HTMLView
            ' Restore raw HTML
            If RichTextBox1.Tag IsNot Nothing AndAlso RichTextBox1.Tag.ToString() <> "" Then
                RichTextBox1.Text = RichTextBox1.Tag.ToString()
                HighLigth_TrackBlock_in_RichTextBox(Start_String.Text, Stop_String.Text)
            End If
        End If
    End Sub

    Private Sub Toggle_SplitOrientation_Click(sender As Object, e As EventArgs) Handles Toggle_SplitOrientation.Click
        b_HorizontalLayout = Not b_HorizontalLayout
        ApplyLayoutOrientation()
    End Sub

    Private Sub ApplyLayoutOrientation()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()

        If b_HorizontalLayout Then
            ' Switch to horizontal (top/bottom) split
            Me.SplitContainer1.Orientation = Orientation.Horizontal
            ' Restore a sensible splitter distance for horizontal mode
            Dim newDist As Integer = CInt(Me.SplitContainer1.Height * 0.5)
            If newDist < 200 Then newDist = 200
            Try
                Me.SplitContainer1.SplitterDistance = newDist
            Catch
            End Try

            ' --- Reposition Panel2 controls for horizontal layout ---
            ' Left column (fixed ~400px): all labels + fields + action buttons
            ' Right column (fills rest): RichTextBox1

            ' Label1 - Track URL label
            Label1.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Label1.Location = New Point(0, 0)
            Label1.Size = New Size(395, 18)

            ' Track_URL textbox
            Track_URL.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Track_URL.Location = New Point(0, 19)
            Track_URL.Size = New Size(265, 27)

            ' Browser_TrackURL button
            Browser_TrackURL.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Browser_TrackURL.Location = New Point(267, 9)

            ' Download_Track_URL button
            Download_Track_URL.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Download_Track_URL.Location = New Point(309, 9)

            ' Save_Track button
            Save_Track.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Save_Track.Location = New Point(351, 9)

            ' Label2 - Start String label
            Label2.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Label2.Location = New Point(0, 55)
            Label2.Size = New Size(395, 18)

            ' Start_String textbox
            Start_String.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Start_String.Location = New Point(0, 74)
            Start_String.Size = New Size(352, 27)

            ' Go_To_Start_String button
            Go_To_Start_String.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Go_To_Start_String.Location = New Point(354, 64)

            ' Label3 - Stop String label
            Label3.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Label3.Location = New Point(0, 110)
            Label3.Size = New Size(395, 18)

            ' Stop_String textbox
            Stop_String.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Stop_String.Location = New Point(0, 129)
            Stop_String.Size = New Size(352, 27)

            ' Go_To_Stop_String button
            Go_To_Stop_String.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Go_To_Stop_String.Location = New Point(354, 119)

            ' Label4 - Search label
            Label4.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Label4.Location = New Point(0, 165)
            Label4.Size = New Size(395, 18)

            ' Find_String textbox
            Find_String.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Find_String.Location = New Point(0, 184)
            Find_String.Size = New Size(185, 27)

            ' Search buttons row
            Search_From_Top.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Search_From_Top.Location = New Point(187, 174)

            Search_From_CARET.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Search_From_CARET.Location = New Point(229, 174)

            Reverse_From_CARET.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Reverse_From_CARET.Location = New Point(271, 174)

            Reverse_From_BOTTOM.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Reverse_From_BOTTOM.Location = New Point(313, 174)

            ' Label_FindCount below search row
            Label_FindCount.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            Label_FindCount.Location = New Point(0, 216)
            Label_FindCount.Size = New Size(395, 18)

            ' RichTextBox1 fills the right side
            RichTextBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            RichTextBox1.Location = New Point(400, 0)
            RichTextBox1.Size = New Size(Math.Max(100, Me.SplitContainer1.Panel2.ClientSize.Width - 405), Me.SplitContainer1.Panel2.ClientSize.Height)

        Else
            ' Switch back to vertical (left/right) split - restore original positions
            Me.SplitContainer1.Orientation = Orientation.Vertical
            ' Restore splitter distance for vertical mode
            Dim newDist As Integer = CInt(Me.SplitContainer1.Width * 0.45)
            If newDist < 200 Then newDist = 200
            Try
                Me.SplitContainer1.SplitterDistance = newDist
            Catch
            End Try

            ' Restore all Panel2 controls to original Designer positions

            ' Label1
            Label1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            Label1.Location = New Point(5, 0)
            Label1.Size = New Size(935, 55)

            ' Track_URL
            Track_URL.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            Track_URL.Location = New Point(15, 22)
            Track_URL.Size = New Size(775, 27)

            ' Browser_TrackURL
            Browser_TrackURL.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            Browser_TrackURL.Location = New Point(800, 8)

            ' Download_Track_URL
            Download_Track_URL.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            Download_Track_URL.Location = New Point(845, 8)

            ' Save_Track
            Save_Track.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            Save_Track.Location = New Point(890, 8)

            ' Label2
            Label2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            Label2.Location = New Point(5, 60)
            Label2.Size = New Size(935, 55)

            ' Start_String
            Start_String.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            Start_String.Location = New Point(15, 82)
            Start_String.Size = New Size(865, 27)

            ' Go_To_Start_String
            Go_To_Start_String.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            Go_To_Start_String.Location = New Point(890, 68)

            ' Label3
            Label3.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            Label3.Location = New Point(5, 120)
            Label3.Size = New Size(935, 55)

            ' Stop_String
            Stop_String.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            Stop_String.Location = New Point(15, 143)
            Stop_String.Size = New Size(865, 27)

            ' Go_To_Stop_String
            Go_To_Stop_String.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            Go_To_Stop_String.Location = New Point(890, 128)

            ' RichTextBox1
            RichTextBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            RichTextBox1.Location = New Point(5, 185)
            RichTextBox1.Size = New Size(935, 440)

            ' Label4
            Label4.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            Label4.Location = New Point(5, 635)
            Label4.Size = New Size(935, 55)

            ' Label_FindCount
            Label_FindCount.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            Label_FindCount.Location = New Point(62, 637)
            Label_FindCount.Size = New Size(683, 18)

            ' Find_String
            Find_String.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            Find_String.Location = New Point(15, 658)
            Find_String.Size = New Size(730, 27)

            ' Search_From_Top
            Search_From_Top.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            Search_From_Top.Location = New Point(755, 643)

            ' Search_From_CARET
            Search_From_CARET.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            Search_From_CARET.Location = New Point(800, 643)

            ' Reverse_From_CARET
            Reverse_From_CARET.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            Reverse_From_CARET.Location = New Point(845, 643)

            ' Reverse_From_BOTTOM
            Reverse_From_BOTTOM.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            Reverse_From_BOTTOM.Location = New Point(890, 643)
        End If

        Me.SplitContainer1.Panel2.ResumeLayout(True)
        Me.SplitContainer1.ResumeLayout(True)
    End Sub
    
    Private Function StripHtmlTags(html As String) As String
        ' Remove script and style blocks entirely
        Dim cleaned As String = System.Text.RegularExpressions.Regex.Replace(html, "<script[^>]*>[\s\S]*?</script>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "<style[^>]*>[\s\S]*?</style>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        ' Remove all HTML tags
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "<[^>]+>", "")
        ' Decode common HTML entities
        cleaned = cleaned.Replace("&amp;", "&")
        cleaned = cleaned.Replace("&lt;", "<")
        cleaned = cleaned.Replace("&gt;", ">")
        cleaned = cleaned.Replace("&quot;", """")
        cleaned = cleaned.Replace("&nbsp;", " ")
        cleaned = cleaned.Replace("&#39;", "'")
        ' Decode all numeric HTML entities (&#9658; &#169; etc.)
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "&#(\d+);", Function(m)
            Try
                Return ChrW(Integer.Parse(m.Groups(1).Value))
            Catch
                Return m.Value
            End Try
        End Function)
        ' Decode hex HTML entities (&#x25BA; etc.)
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "&#x([0-9A-Fa-f]+);", Function(m)
            Try
                Return ChrW(Integer.Parse(m.Groups(1).Value, Globalization.NumberStyles.HexNumber))
            Catch
                Return m.Value
            End Try
        End Function)
        ' Collapse multiple blank lines into one
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "(\r?\n\s*){3,}", vbCrLf & vbCrLf)
        Return cleaned.Trim()
    End Function
    
    '======================================
    ' INTEGRATED Form2 CODE - ENDS HERE
    '======================================
    
    Private Function GetAllSuitePaths() As List(Of String)
        Dim suitePaths As New List(Of String)
        Dim parentPath As String = Path.GetDirectoryName(s_SyMenuSuite_Path)
        
        ' Always include the default SyMenuSuite
        suitePaths.Add(s_SyMenuSuite_Path)
        
        ' Find all other suite folders that have _Cache subfolder
        If Directory.Exists(parentPath) Then
            For Each suiteFolder In Directory.GetDirectories(parentPath)
                If Directory.Exists(Path.Combine(suiteFolder, "_Cache")) Then
                    If suiteFolder <> s_SyMenuSuite_Path Then
                        suitePaths.Add(suiteFolder)
                    End If
                End If
            Next
        End If
        
        Return suitePaths
    End Function
    
    '======================================
    ' Rest of Form1 code (menu handlers, etc.)
    '======================================
    Private Function FuzzyMatch(searchTerm As String, publisherName As String) As Boolean
        ' Normalize both strings: lowercase and remove spaces
        Dim search As String = searchTerm.Trim().ToLower().Replace(" ", "")
        Dim publisher As String = publisherName.ToLower().Replace(" ", "")

        ' Exact containment match (ignoring case and spaces)
        If publisher.Contains(search) OrElse search.Contains(publisher) Then
            Return True
        End If

        ' Skip fuzzy matching for very short search terms
        If search.Length < 3 Then Return False

        ' Fuzzy match: calculate how many characters match in sequence
        Dim matchCount As Integer = 0
        Dim pubIndex As Integer = 0
        For Each c As Char In search
            For j As Integer = pubIndex To publisher.Length - 1
                If publisher(j) = c Then
                    matchCount += 1
                    pubIndex = j + 1
                    Exit For
                End If
            Next
        Next

        ' Consider it a match if 75% or more of the search characters were found in order
        Dim matchRatio As Double = matchCount / search.Length
        Return matchRatio >= 0.75
    End Function
    
    Private Sub SPS_P_Tracker_Name_KeyDown(sender As Object, e As KeyEventArgs) Handles SPS_P_Tracker_Name.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ReBuild_SPS_List.PerformClick()
        End If
    End Sub
    
    Private Sub ReBuild_SPS_List_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReBuild_SPS_List.Click
        ReBuild_SPS_List.Enabled = False
        ' Check for unsaved changes before rebuilding
        If SPS_P_ListFile_Name.BackColor.Equals(If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)) Then
            Dim result As Integer
            result = ShowThemedMessageBox("You have unsaved changes. Save before rebuilding?", "Save Track List", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
            Select Case result
                Case DialogResult.Yes
                    SavePATListFile(s_SPS_P_ListFile_Path)
                Case DialogResult.No
                    ' Continue without saving
                Case DialogResult.Cancel
                    Return
            End Select
        End If
        
        ' Warn if search bar is empty (will load ALL SPS files)
        If SPS_P_Tracker_Name.Text.Trim() = "" Then
            Dim result As Integer
            result = ShowThemedMessageBox("The search bar is empty. This will load ALL SPS files from all suites," & vbCrLf & "which may take a long time if there are many files." & vbCrLf & vbCrLf & "Do you want to continue?", "Load All SPS Files", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.No Then
                Return
            End If
        End If

        SPS_P_ListView.BeginUpdate()
        SPS_P_ImageList.Images.Clear()
        SPS_P_ListView.Items.Clear()
        
        ' Load PAT data from the saved file on disk
        Dim savedPATData As New Dictionary(Of String, Dictionary(Of String, String))
        If File.Exists(s_SPS_P_ListFile_Path) Then
            Dim s_PAT_text As String = File.ReadAllText(s_SPS_P_ListFile_Path, Encoding.UTF8)
            Dim idx As Integer = 1
            Dim s_entry As String = S_GetsNBlockFromText(s_PAT_text, "<SPSPublishedEntry>", "</SPSPublishedEntry>", idx)
            While s_entry <> ""
                Dim name As String = S_GetsNBlockFromText(s_entry, "<ProgramName>", "</ProgramName>", 1)
                If name <> "" AndAlso Not savedPATData.ContainsKey(name) Then
                    Dim fields As New Dictionary(Of String, String)
                    fields("TrackURL") = S_GetsNBlockFromText(s_entry, "<TrackURL>", "</TrackURL>", 1)
                    fields("TrackStartString") = S_GetsNBlockFromText(s_entry, "<TrackStartString>", "</TrackStartString>", 1)
                    fields("TrackStopString") = S_GetsNBlockFromText(s_entry, "<TrackStopString>", "</TrackStopString>", 1)
                    fields("TrackBlockHash") = S_GetsNBlockFromText(s_entry, "<TrackBlockHash>", "</TrackBlockHash>", 1)
                    fields("LatestVersion") = S_GetsNBlockFromText(s_entry, "<LatestVersion>", "</LatestVersion>", 1)
                    savedPATData(name) = fields
                End If
                idx += 1
                s_entry = S_GetsNBlockFromText(s_PAT_text, "<SPSPublishedEntry>", "</SPSPublishedEntry>", idx)
            End While
        End If
        
        ' Process the list of files found in the directory.
        Dim s_fileName, s_SPS_File_text, s_SPS_PublisherName, s_SPS_Name, s_SPS_ProgramPublisherWebSite,
            s_SPS_Version, s_SPS_ReleaseDate, s_SPS_DownloadUrl, s_SPS_DownloadSizeKb, s_SPS_ProgramIconBase64,
            s_SPS_CreationDate, s_SPS_ModificationDate As String
        Dim img_SPS_Icon As System.Drawing.Image
        Dim thisFile As System.IO.FileInfo
        Dim AddSPS As Boolean
        Dim i_progress As Integer = 1
        'Build the SPS Suite files
        Dim allSuitePaths As List(Of String) = GetAllSuitePaths()
        Dim s_fileEntries As String() = {} ' Start empty
        ProgressBar1.Value = 20
        Dim shellType As Type = Type.GetTypeFromProgID("Shell.Application", True)
        Dim shellObj As Object = Activator.CreateInstance(shellType)
        
        ' For each suite, extract zips into its own _TmpPAT and scan for .sps files
        For Each currentSuitePath In allSuitePaths
            Dim cachePath As String = Path.Combine(currentSuitePath, "_Cache")
            If Not Directory.Exists(cachePath) Then Continue For
            
            Dim tmpPath As String = Path.Combine(currentSuitePath, "_Trash", "_TmpPAT")
            Dim cacheFiles As String() = Directory.GetFiles(cachePath)
            Dim hasZips As Boolean = cacheFiles.Any(Function(f) Path.GetExtension(f).ToLower() = ".zip")
            
            Dim scanFolder As String
            If hasZips Then
                ' This suite has zip files - extract them to a temp folder
                If IO.Directory.Exists(tmpPath) Then
                    My.Computer.FileSystem.DeleteDirectory(tmpPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
                End If
                If Not IO.Directory.Exists(Path.Combine(currentSuitePath, "_Trash")) Then
                    IO.Directory.CreateDirectory(Path.Combine(currentSuitePath, "_Trash"))
                End If
                IO.Directory.CreateDirectory(tmpPath)
                ' Extract zips
                For Each s_fileName In cacheFiles
                    thisFile = My.Computer.FileSystem.GetFileInfo(s_fileName)
                    If thisFile.Extension = ".zip" Then
                        Dim outputFolder As Object = shellType.InvokeMember("NameSpace", BindingFlags.InvokeMethod, Nothing, shellObj, New [Object]() {tmpPath})
                        Dim inputZipFile As Object = shellType.InvokeMember("NameSpace", BindingFlags.InvokeMethod, Nothing, shellObj, New [Object]() {thisFile.FullName})
                        outputFolder.CopyHere((inputZipFile.Items), 4)
                    End If
                Next
                ' Also copy any loose .sps files from _Cache
                For Each s_fileName In cacheFiles
                    thisFile = My.Computer.FileSystem.GetFileInfo(s_fileName)
                    If thisFile.Extension = ".sps" Then
                        IO.File.Copy(s_fileName, Path.Combine(tmpPath, thisFile.Name), True)
                    End If
                Next
                scanFolder = tmpPath
            Else
                ' No zips - scan _Cache directly for .sps files
                scanFolder = cachePath
            End If
                If Directory.Exists(scanFolder) Then
                s_fileEntries = Directory.GetFiles(scanFolder)
                For Each s_fileName In s_fileEntries
                    Dim barValue As Integer = CInt(50 + i_progress * (ProgressBar1.Maximum - 50) / s_fileEntries.Length)
                    If barValue < ProgressBar1.Minimum Then barValue = ProgressBar1.Minimum
                    If barValue > ProgressBar1.Maximum Then barValue = ProgressBar1.Maximum
                    ProgressBar1.Value = barValue
                    i_progress = i_progress + 1
                    thisFile = My.Computer.FileSystem.GetFileInfo(s_fileName)
                    
                    If thisFile.Extension = ".sps" Then
                        s_SPS_File_text = File.ReadAllText(s_fileName, Encoding.UTF8)
                        s_SPS_PublisherName = S_GetsNBlockFromText(s_SPS_File_text, "<SPSPublisherName>", "</SPSPublisherName>", 1)
                        Dim b_MatchPublisher As Boolean = False
                        If SPS_P_Tracker_Name.Text.Trim() = "" Then
                            b_MatchPublisher = True
                        ElseIf s_SPS_PublisherName <> "" Then
                            ' Fuzzy match: ignores case, spaces, and tolerates minor typos
                            b_MatchPublisher = FuzzyMatch(SPS_P_Tracker_Name.Text, s_SPS_PublisherName)
                        End If
                        If b_MatchPublisher Then
                            s_SPS_Name = S_GetsNBlockFromText(s_SPS_File_text, "<ProgramName>", "</ProgramName>", 1)
                            s_SPS_ProgramPublisherWebSite = S_GetsNBlockFromText(s_SPS_File_text, "<ProgramPublisherWebSite>", "</ProgramPublisherWebSite>", 1)
                            s_SPS_Version = S_GetsNBlockFromText(s_SPS_File_text, "<Version>", "</Version>", 1)
                            s_SPS_ReleaseDate = S_GetsNBlockFromText(s_SPS_File_text, "<ReleaseDate>", "</ReleaseDate>", 1)
                            s_SPS_DownloadUrl = S_GetsNBlockFromText(s_SPS_File_text, "<DownloadUrl>", "</DownloadUrl>", 1)
                            s_SPS_DownloadSizeKb = S_GetsNBlockFromText(s_SPS_File_text, "<DownloadSizeKb>", "</DownloadSizeKb>", 1)
                            s_SPS_ProgramIconBase64 = S_GetsNBlockFromText(s_SPS_File_text, "<ProgramIconBase64>", "</ProgramIconBase64>", 1)
                            s_SPS_CreationDate = thisFile.CreationTime.ToString("yyyy-MM-dd")
                            s_SPS_ModificationDate = thisFile.LastWriteTime.ToString("yyyy-MM-dd")
                            'Get the app icon (16x16)
                            Try
                                Dim imageBytes() As Byte = System.Convert.FromBase64String(s_SPS_ProgramIconBase64)
                                Using stream = New System.IO.MemoryStream(imageBytes, 0, imageBytes.Length)
                                    img_SPS_Icon = System.Drawing.Image.FromStream(stream)
                                End Using
                            Catch
                                img_SPS_Icon = New Bitmap(16, 16)
                            End Try
                            'Inicialize the app as new and search in the actual SPS List
                            AddSPS = True
                            For i = 0 To SPS_P_ListView.Items.Count - 1
                                If SPS_P_ListView.Items(i).SubItems(c_SPS_Name).Text = s_SPS_Name Then
                                    AddSPS = False
                                    SPS_P_ListView.Items(i).SubItems(c_SPS_Name).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
                                    SPS_P_ListView.Items(i).UseItemStyleForSubItems = False
                                    SPS_P_ListView.Items(i).SubItems(c_SPS_Name).Text = s_SPS_Name
                                    SPS_P_ListView.Items(i).SubItems(c_Version).Text = s_SPS_Version
                                    SPS_P_ListView.Items(i).SubItems(c_ReleaseDate).Text = s_SPS_ReleaseDate
                                    SPS_P_ListView.Items(i).SubItems(c_DownloadUrl).Text = s_SPS_DownloadUrl
                                    SPS_P_ListView.Items(i).SubItems(c_DownloadUrl).BackColor = Color.Empty
                                    SPS_P_ListView.Items(i).SubItems(c_DownloadSizeKb).Text = s_SPS_DownloadSizeKb
                                    SPS_P_ListView.Items(i).SubItems(c_DownloadSizeKb).BackColor = Color.Empty
                                    SPS_P_ListView.Items(i).SubItems(c_SPSCreationDate).Text = s_SPS_CreationDate
                                    SPS_P_ListView.Items(i).SubItems(c_SPSModificationDate).Text = s_SPS_ModificationDate
                                    SPS_P_ListView.Items(i).SubItems(c_SPSPublisherName).Text = s_SPS_PublisherName
                                    SPS_P_ListView.Items(i).SubItems(c_SuiteName).Text = Path.GetFileName(currentSuitePath)
                                    SPS_P_ListView.Items(i).Tag = currentSuitePath
                                    SPS_P_ImageList.Images.Add(s_SPS_Name, img_SPS_Icon)
                                    SPS_P_ListView.Items(i).ImageKey = s_SPS_Name
                                    Exit For
                                End If
                            Next
                            If AddSPS = True Then 'New SPS
                                Dim lvi_newItem As New ListViewItem
                                lvi_newItem.SubItems(0).Text = s_SPS_Name
                                lvi_newItem.SubItems(0).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
                                lvi_newItem.UseItemStyleForSubItems = False
                                If savedPATData.ContainsKey(s_SPS_Name) Then
                                    Dim saved = savedPATData(s_SPS_Name)
                                    lvi_newItem.SubItems.Add(saved("TrackURL"))
                                    lvi_newItem.SubItems.Add(saved("TrackStartString"))
                                    lvi_newItem.SubItems.Add(saved("TrackStopString"))
                                    lvi_newItem.SubItems.Add(saved("TrackBlockHash"))
                                    lvi_newItem.SubItems.Add(s_SPS_Version)
                                    lvi_newItem.SubItems.Add(saved("LatestVersion"))
                                Else
                                    lvi_newItem.SubItems.Add(s_SPS_ProgramPublisherWebSite)
                                    lvi_newItem.SubItems.Add("")
                                    lvi_newItem.SubItems.Add("")
                                    lvi_newItem.SubItems.Add("")
                                    lvi_newItem.SubItems.Add(s_SPS_Version)
                                    lvi_newItem.SubItems.Add("")
                                End If
                                lvi_newItem.SubItems.Add(s_SPS_ReleaseDate)
                                lvi_newItem.SubItems.Add(s_SPS_DownloadUrl)
                                lvi_newItem.SubItems.Add(s_SPS_DownloadSizeKb)
                                lvi_newItem.SubItems.Add(s_SPS_CreationDate)
                                lvi_newItem.SubItems.Add(s_SPS_ModificationDate)
                                lvi_newItem.SubItems.Add(s_SPS_PublisherName)
                                lvi_newItem.SubItems.Add(Path.GetFileName(currentSuitePath))
                                lvi_newItem.Tag = currentSuitePath
                                lvi_newItem.ImageKey = s_SPS_Name
                                SPS_P_ImageList.Images.Add(s_SPS_Name, img_SPS_Icon)
                                SPS_P_ListView.Items.Add(lvi_newItem)
                            End If
                        End If
                    End If
                Next s_fileName
            End If
        Next currentSuitePath
        'Delete the _TmpPAT folders for all suites
        For Each suitePath In allSuitePaths
            Dim tmpCleanup As String = Path.Combine(suitePath, "_Trash", "_TmpPAT")
            If IO.Directory.Exists(tmpCleanup) Then
                My.Computer.FileSystem.DeleteDirectory(tmpCleanup, FileIO.DeleteDirectoryOption.DeleteAllContents)
            End If
        Next
        
        'Remove Multiple Selected Items 
        For Each item As ListViewItem In SPS_P_ListView.SelectedItems
            item.Remove()
        Next
        'It must order the ListView in order to match with the ImageView added icons
        SPS_P_ListView.Refresh()
        SPS_P_Tracker_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
        SPS_P_Tracker_Name.Refresh()
        SPS_P_ListFile_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
        File_Save.Enabled = True
        File_SaveAs.Enabled = True
        SPS_P_ListView.EndUpdate()
        SPS_P_Tracker_Name.Refresh()
        SPS_P_Publisher_Number.Text = SPS_P_ListView.Items.Count & " SPS Files"
        SPS_P_Publisher_Number.Refresh()
        SelectAll_CheckBox.Checked = False
        ReBuild_SPS_List.Enabled = True
    End Sub
        
    Private Sub SelectAll_CheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAll_CheckBox.CheckedChanged
        For i = 0 To SPS_P_ListView.Items.Count - 1
            SPS_P_ListView.Items(i).Checked = SelectAll_CheckBox.Checked()
        Next
    End Sub
    
    Private Sub Track_Selected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Track_Checked.Click
        If bgWorker.IsBusy Then
            ' Already running - act as Stop button
            bgWorker.CancelAsync()
            Track_Checked.Enabled = False
            Return
        End If
        
        ' Check for unsaved changes before tracking
        If SPS_P_ListFile_Name.BackColor.Equals(If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)) Then
            Dim result As Integer
            result = ShowThemedMessageBox("You have unsaved changes. Save before checking tracks?", "Save Track List", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
            Select Case result
                Case DialogResult.Yes
                    SavePATListFile(s_SPS_P_ListFile_Path)
                Case DialogResult.No
                    ' Continue without saving
                Case DialogResult.Cancel
                    Return
            End Select
        End If
                
        Track_Checked.Image = New Bitmap(ToolStrip_DeleteSelectedTrack.Image, New Size(20, 20))
        ReBuild_SPS_List.Enabled = False
        SelectAll_CheckBox.Visible = False
        LoadToolStripMenuItem.Enabled = False
        SaveToolStripMenuItem.Enabled = False
        SaveAsToolStripMenuItem.Enabled = False
        bgWorker.RunWorkerAsync()
    End Sub

    Private Function CheckSPS_Download(ByVal trackURL As String, ByVal startString As String, ByVal stopString As String, ByVal downloadUrl As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        result("TrackSuccess") = "False"
        result("DownloadSuccess") = "False"
        result("WebBlock") = ""
        result("WebBlockHash") = ""
        result("LatestVersion") = ""
        result("RemoteSizeKb") = "0"
        
        ' Download track page
        Try
            Dim tempCookies As New CookieContainer
            Dim wrRequest As HttpWebRequest = CType(HttpWebRequest.Create(trackURL), HttpWebRequest)
            ' If any of your tracked URLs start failing because the server rejects your User-Agent as outdated or suspicious, you can go to that site, grab a current User-Agent string, and update the one in your code.
            wrRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0"
            wrRequest.Timeout = 10000
            wrRequest.CookieContainer = tempCookies
            Dim wrResponse As HttpWebResponse = CType(wrRequest.GetResponse(), HttpWebResponse)
            Dim s_Web_Page As String
            Using sr As New StreamReader(wrResponse.GetResponseStream())
                s_Web_Page = sr.ReadToEnd()
                sr.Close()
                wrResponse.Close()
            End Using
            
            Dim s_Web_Block As String = S_GetsNBlockFromText(s_Web_Page, startString, stopString, 1)
            If s_Web_Block <> "" Then
                result("LatestVersion") = s_Web_Block.Trim()
            End If
            If s_Web_Block = "" Then s_Web_Block = s_Web_Page
            
            Using Md5 As New System.Security.Cryptography.MD5CryptoServiceProvider()
                Dim ueCodigo As New System.Text.UnicodeEncoding()
                Dim bHash() As Byte = Md5.ComputeHash(ueCodigo.GetBytes(s_Web_Block))
                result("WebBlockHash") = Convert.ToBase64String(bHash)
            End Using
            result("WebBlock") = s_Web_Block
            result("TrackSuccess") = "True"
        Catch ex As Exception
            result("TrackSuccess") = "False"
        End Try
        
        ' Check download size
        Try
            Dim tempCookies As New CookieContainer
            Dim wrRequest As HttpWebRequest = CType(HttpWebRequest.Create(downloadUrl), HttpWebRequest)
            wrRequest.UserAgent = "Wget/1.25.0"
            wrRequest.Timeout = 10000
            Using wrResponse As HttpWebResponse = CType(wrRequest.GetResponse(), HttpWebResponse)
                result("RemoteSizeKb") = CLng(Math.Round(wrResponse.ContentLength / 1024, 0)).ToString()
                wrResponse.Close()
            End Using
            result("DownloadSuccess") = "True"
        Catch ex As Exception
            result("DownloadSuccess") = "False"
        End Try
        
        Return result
    End Function
    
    Public Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'Check if track list needs saving first
        If SPS_P_ListFile_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed) Then
            Dim result As Integer = ShowThemedMessageBox("Save the modificated Track List?", "Save Track List", MessageBoxButtons.YesNoCancel)
            Select Case result
                Case DialogResult.Yes
                    SaveToolStripMenuItem_Click(sender, e)
                    e.Cancel = True
                    Return
                Case DialogResult.No
                    e.Cancel = False
                Case DialogResult.Cancel
                    e.Cancel = True
                    Return
            End Select
        End If

        'Auto-save config on exit - only if paths are valid
        If Not String.IsNullOrEmpty(s_ConfigPAT_FilePath) Then
            s_ConfigPAT_text = "<PAT_ListFile>" & Path.GetFileName(s_SPS_P_ListFile_Path) & "</PAT_ListFile>" & vbCrLf
            s_ConfigPAT_text &= "<Form1_x>" & Me.Left & "</Form1_x>" & vbCrLf
            s_ConfigPAT_text &= "<Form1_y>" & Me.Top & "</Form1_y>" & vbCrLf
            s_ConfigPAT_text &= "<Form1_w>" & Me.Width & "</Form1_w>" & vbCrLf
            s_ConfigPAT_text &= "<Form1_h>" & Me.Height & "</Form1_h>" & vbCrLf
            s_ConfigPAT_text &= "<SplitterDistance>" & Me.SplitContainer1.SplitterDistance & "</SplitterDistance>" & vbCrLf
            If SPS_P_ListView.Columns.Count > 0 Then
                s_ConfigPAT_text &= "<SPS_Name_w>" & Me.SPS_P_ListView.Columns(c_SPS_Name).Width & "</SPS_Name_w>" & vbCrLf
                s_ConfigPAT_text &= "<TrackURL_w>" & Me.SPS_P_ListView.Columns(c_TrackURL).Width & "</TrackURL_w>" & vbCrLf
                s_ConfigPAT_text &= "<TrackStartString_w>" & Me.SPS_P_ListView.Columns(c_TrackStartString).Width & "</TrackStartString_w>" & vbCrLf
                s_ConfigPAT_text &= "<TrackStopString_w>" & Me.SPS_P_ListView.Columns(c_TrackStopString).Width & "</TrackStopString_w>" & vbCrLf
                s_ConfigPAT_text &= "<TrackBlockHash_w>" & Me.SPS_P_ListView.Columns(c_TrackBlockHash).Width & "</TrackBlockHash_w>" & vbCrLf
                s_ConfigPAT_text &= "<Version_w>" & Me.SPS_P_ListView.Columns(c_Version).Width & "</Version_w>" & vbCrLf
                s_ConfigPAT_text &= "<LatestVersion_w>" & Me.SPS_P_ListView.Columns(c_LatestVersion).Width & "</LatestVersion_w>" & vbCrLf
                s_ConfigPAT_text &= "<ReleaseDate_w>" & Me.SPS_P_ListView.Columns(c_ReleaseDate).Width & "</ReleaseDate_w>" & vbCrLf
                s_ConfigPAT_text &= "<DownloadUrl_w>" & Me.SPS_P_ListView.Columns(c_DownloadUrl).Width & "</DownloadUrl_w>" & vbCrLf
                s_ConfigPAT_text &= "<DownloadSizeKb_w>" & Me.SPS_P_ListView.Columns(c_DownloadSizeKb).Width & "</DownloadSizeKb_w>" & vbCrLf
                s_ConfigPAT_text &= "<SPSCreationDate_w>" & Me.SPS_P_ListView.Columns(c_SPSCreationDate).Width & "</SPSCreationDate_w>" & vbCrLf
                s_ConfigPAT_text &= "<SPSModificationDate_w>" & Me.SPS_P_ListView.Columns(c_SPSModificationDate).Width & "</SPSModificationDate_w>" & vbCrLf
                s_ConfigPAT_text &= "<SPSPublisherName_w>" & Me.SPS_P_ListView.Columns(c_SPSPublisherName).Width & "</SPSPublisherName_w>" & vbCrLf
                s_ConfigPAT_text &= "<SuiteName_w>" & Me.SPS_P_ListView.Columns(c_SuiteName).Width & "</SuiteName_w>" & vbCrLf
                Dim colOrder As String = ""
                For i As Integer = 0 To SPS_P_ListView.Columns.Count - 1
                    If i > 0 Then colOrder &= ","
                    colOrder &= SPS_P_ListView.Columns(i).DisplayIndex.ToString()
                Next
                s_ConfigPAT_text &= "<ColumnOrder>" & colOrder & "</ColumnOrder>" & vbCrLf
            End If
            s_ConfigPAT_text &= "<SyMenuSuitePath>" & s_SyMenuSuite_Path & "</SyMenuSuitePath>" & vbCrLf
            s_ConfigPAT_text &= "<DarkMode>" & DarkModeToolStripMenuItem.Checked.ToString() & "</DarkMode>" & vbCrLf
            s_ConfigPAT_text &= "<SplitOrientation>" & If(b_HorizontalLayout, "Horizontal", "Vertical") & "</SplitOrientation>" & vbCrLf
            File.WriteAllText(s_ConfigPAT_FilePath, s_ConfigPAT_text, Encoding.UTF8)
        End If

        'Delete the _TmpPAT folders if they exist
        Try
            For Each suitePath In GetAllSuitePaths()
                Dim tmpCleanup As String = Path.Combine(suitePath, "_Trash", "_TmpPAT")
                If IO.Directory.Exists(tmpCleanup) Then
                    My.Computer.FileSystem.DeleteDirectory(tmpCleanup, FileIO.DeleteDirectoryOption.DeleteAllContents)
                End If
            Next
        Catch
        End Try
    End Sub
        
    Private Sub Help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Help.Click
        Try
            Dim bgColor As Color = If(DarkModeToolStripMenuItem.Checked, m_DarkModeBackground30, Color.White)
            Dim hf As New HelpForm(DarkModeToolStripMenuItem.Checked, bgColor)
            hf.ShowDialog(Me)
        Catch ex As Exception
            ShowThemedMessageBox("Error: " & ex.Message & vbCrLf & ex.StackTrace, "Help Error")
        End Try
    End Sub
  
    Private Sub SPS_P_ListView_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles SPS_P_ListView.ColumnClick
        If e.Column = i_SortColumn Then
            bln_SortAscending = Not bln_SortAscending
        Else
            ' New column - detect if already sorted ascending, if so start descending
            Dim alreadyAscending As Boolean = True
            Dim comparer As New ListViewItemComparer(e.Column, True)
            For i As Integer = 0 To SPS_P_ListView.Items.Count - 2
                If comparer.Compare(SPS_P_ListView.Items(i), SPS_P_ListView.Items(i + 1)) > 0 Then
                    alreadyAscending = False
                    Exit For
                End If
            Next
            i_SortColumn = e.Column
            bln_SortAscending = If(alreadyAscending, False, True)
        End If

        ' Clone all items
        Dim count As Integer = SPS_P_ListView.Items.Count
        If count = 0 Then Return
        
        Dim clones(count - 1) As ListViewItem
        For i As Integer = 0 To count - 1
            clones(i) = CType(SPS_P_ListView.Items(i).Clone(), ListViewItem)
        Next
        
        Array.Sort(clones, New ListViewItemComparer(e.Column, bln_SortAscending))
        
        ' Copy sorted data back into existing rows
        SPS_P_ListView.BeginUpdate()
        For i As Integer = 0 To count - 1
            Dim src As ListViewItem = clones(i)
            Dim dest As ListViewItem = SPS_P_ListView.Items(i)
            dest.Text = src.Text
            dest.ImageKey = src.ImageKey
            dest.Tag = src.Tag
            dest.Checked = src.Checked
            dest.UseItemStyleForSubItems = src.UseItemStyleForSubItems
            dest.SubItems(0).BackColor = src.SubItems(0).BackColor
            For j As Integer = 1 To src.SubItems.Count - 1
                dest.SubItems(j).Text = src.SubItems(j).Text
                dest.SubItems(j).BackColor = src.SubItems(j).BackColor
                dest.SubItems(j).ForeColor = src.SubItems(j).ForeColor
            Next
        Next
        SPS_P_ListView.EndUpdate()
        
        ' Update column headers to show sort direction (right-aligned arrow)
        For i As Integer = 0 To SPS_P_ListView.Columns.Count - 1
            Dim header As String = SPS_P_ListView.Columns(i).Text
            ' Remove any existing arrow from end
            If header.EndsWith(" ▲") Then
                header = header.Substring(0, header.Length - 2)
            ElseIf header.EndsWith(" ▼") Then
                header = header.Substring(0, header.Length - 2)
            End If
            If i = e.Column Then
                SPS_P_ListView.Columns(i).Text = header & If(bln_SortAscending, " ▲", " ▼")
            Else
                SPS_P_ListView.Columns(i).Text = header
            End If
        Next
    End Sub
    
    'MenuStrip Subrutines
    Private Sub LoadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadToolStripMenuItem.Click
        OpenFileDialog1.FileName = Path.GetFileName(s_SPS_P_ListFile_Path)
        OpenFileDialog1.InitialDirectory = Path.GetDirectoryName(s_SPS_P_ListFile_Path)
        OpenFileDialog1.DefaultExt = ".xml"
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Filter = "Text Files|*.xml"
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.AutoUpgradeEnabled = True
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        If result = DialogResult.OK Then
            s_SPS_P_ListFile_Path = OpenFileDialog1.FileName
            SPS_P_ListFile_Name.BackColor = Color.Empty
            SPS_P_ListFile_Name.ForeColor = If(DarkModeToolStripMenuItem.Checked, Color.White, Color.Black)
            SPS_P_ListFile_Name.Text = Path.GetFileName(s_SPS_P_ListFile_Path)
            SPS_P_ListFile_Name.Refresh()
            Charge_SPS_P_ListView()
            File_Save.Enabled = False
            File_SaveAs.Enabled = False
            SelectAll_CheckBox.Checked = False
        End If
    End Sub
    
    Private Sub Open_File_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Open_File.Click
        LoadToolStripMenuItem_Click(sender, e)
    End Sub
    
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click, File_Save.Click
        SavePATListFile(s_SPS_P_ListFile_Path)
    End Sub
       
    Private Sub SavePATListFile(ByVal s_FilePath As String)
        Dim s_SPS_P_ListFile_text As String = "<SPSTrackerName>" & SPS_P_Tracker_Name.Text & "</SPSTrackerName>" & vbCrLf
        For i = 0 To SPS_P_ListView.Items.Count - 1
            Dim s_SPS_P_IconBase64 As String = ""
            Try
                Using stream As MemoryStream = New System.IO.MemoryStream()
                    SPS_P_ImageList.Images(SPS_P_ImageList.Images.IndexOfKey(SPS_P_ListView.Items(i).SubItems(c_SPS_Name).Text)).Save(stream, System.Drawing.Imaging.ImageFormat.Png)
                    stream.Close()
                    Dim imageBytes() As Byte = stream.ToArray()
                    s_SPS_P_IconBase64 = System.Convert.ToBase64String(imageBytes)
                End Using
            Catch
                s_SPS_P_IconBase64 = ""
            End Try
            s_SPS_P_ListFile_text = s_SPS_P_ListFile_text & "<SPSPublishedEntry>" & vbCrLf &
            vbTab & "<ProgramName>" & SPS_P_ListView.Items(i).SubItems(c_SPS_Name).Text & "</ProgramName>" & vbCrLf &
            vbTab & "<TrackURL>" & SPS_P_ListView.Items(i).SubItems(c_TrackURL).Text & "</TrackURL>" & vbCrLf &
            vbTab & "<TrackStartString>" & SPS_P_ListView.Items(i).SubItems(c_TrackStartString).Text & "</TrackStartString>" & vbCrLf &
            vbTab & "<TrackStopString>" & SPS_P_ListView.Items(i).SubItems(c_TrackStopString).Text & "</TrackStopString>" & vbCrLf &
            vbTab & "<TrackBlockHash>" & SPS_P_ListView.Items(i).SubItems(c_TrackBlockHash).Text & "</TrackBlockHash>" & vbCrLf &
            vbTab & "<Version>" & SPS_P_ListView.Items(i).SubItems(c_Version).Text & "</Version>" & vbCrLf &
            vbTab & "<LatestVersion>" & SPS_P_ListView.Items(i).SubItems(c_LatestVersion).Text & "</LatestVersion>" & vbCrLf &
            vbTab & "<ReleaseDate>" & SPS_P_ListView.Items(i).SubItems(c_ReleaseDate).Text & "</ReleaseDate>" & vbCrLf &
            vbTab & "<DownloadUrl>" & SPS_P_ListView.Items(i).SubItems(c_DownloadUrl).Text & "</DownloadUrl>" & vbCrLf &
            vbTab & "<DownloadSizeKb>" & SPS_P_ListView.Items(i).SubItems(c_DownloadSizeKb).Text & "</DownloadSizeKb>" & vbCrLf &
            vbTab & "<ProgramIconBase64>" & s_SPS_P_IconBase64 & "</ProgramIconBase64>" & vbCrLf &
            vbTab & "<SPSCreationDate>" & SPS_P_ListView.Items(i).SubItems(c_SPSCreationDate).Text & "</SPSCreationDate>" & vbCrLf &
            vbTab & "<SPSModificationDate>" & SPS_P_ListView.Items(i).SubItems(c_SPSModificationDate).Text & "</SPSModificationDate>" & vbCrLf &
            vbTab & "<SPSPublisherName>" & SPS_P_ListView.Items(i).SubItems(c_SPSPublisherName).Text & "</SPSPublisherName>" & vbCrLf &
            vbTab & "<SuiteName>" & SPS_P_ListView.Items(i).SubItems(c_SuiteName).Text & "</SuiteName>" & vbCrLf &
            "</SPSPublishedEntry>" & vbCrLf
        Next
        File.WriteAllText(s_FilePath, s_SPS_P_ListFile_text, Encoding.UTF8)
        If DarkModeToolStripMenuItem.Checked Then
            SPS_P_ListFile_Name.BackColor = m_DarkModeBackground45
            SPS_P_ListFile_Name.ForeColor = Color.White
        Else
            SPS_P_ListFile_Name.BackColor = Color.Empty
            SPS_P_ListFile_Name.ForeColor = Color.Black
        End If
        SPS_P_ListFile_Name.Text = Path.GetFileName(s_SPS_P_ListFile_Path)
        SPS_P_ListFile_Name.Refresh()
        File_Save.Enabled = False
        File_SaveAs.Enabled = False
    End Sub
    
    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        Me.SaveFileDialog1.FileName = Path.GetFileNameWithoutExtension(s_SPS_P_ListFile_Path)
        Me.SaveFileDialog1.InitialDirectory = Path.GetDirectoryName(s_SPS_P_ListFile_Path)
        Me.SaveFileDialog1.DefaultExt = ".xml"
        Me.SaveFileDialog1.AddExtension = True
        Me.SaveFileDialog1.Filter = "Text Files|*.xml"
        Me.SaveFileDialog1.CheckFileExists = False
        Me.SaveFileDialog1.AutoUpgradeEnabled = False
        Dim result As DialogResult = Me.SaveFileDialog1.ShowDialog()
        If result = DialogResult.OK Then
            s_SPS_P_ListFile_Path = SaveFileDialog1.FileName
            SavePATListFile(s_SPS_P_ListFile_Path)
        End If
    End Sub
    
    Private Sub SaveAndExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAndExitToolStripMenuItem.Click
        SavePATListFile(s_SPS_P_ListFile_Path)
        Me.Close()
    End Sub
    
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    
    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        s_ConfigPAT_text = "<PAT_ListFile>" & Path.GetFileName(s_SPS_P_ListFile_Path) & "</PAT_ListFile>" & vbCrLf
        s_ConfigPAT_text &= "<Form1_x>" & Me.Left & "</Form1_x>" & vbCrLf
        s_ConfigPAT_text &= "<Form1_y>" & Me.Top & "</Form1_y>" & vbCrLf
        s_ConfigPAT_text &= "<Form1_w>" & Me.Width & "</Form1_w>" & vbCrLf
        s_ConfigPAT_text &= "<Form1_h>" & Me.Height & "</Form1_h>" & vbCrLf
        s_ConfigPAT_text &= "<SplitterDistance>" & Me.SplitContainer1.SplitterDistance & "</SplitterDistance>" & vbCrLf
        If SPS_P_ListView.Columns.Count > 0 Then
            s_ConfigPAT_text &= "<SPS_Name_w>" & Me.SPS_P_ListView.Columns(c_SPS_Name).Width & "</SPS_Name_w>" & vbCrLf
            s_ConfigPAT_text &= "<TrackURL_w>" & Me.SPS_P_ListView.Columns(c_TrackURL).Width & "</TrackURL_w>" & vbCrLf
            s_ConfigPAT_text &= "<TrackStartString_w>" & Me.SPS_P_ListView.Columns(c_TrackStartString).Width & "</TrackStartString_w>" & vbCrLf
            s_ConfigPAT_text &= "<TrackStopString_w>" & Me.SPS_P_ListView.Columns(c_TrackStopString).Width & "</TrackStopString_w>" & vbCrLf
            s_ConfigPAT_text &= "<TrackBlockHash_w>" & Me.SPS_P_ListView.Columns(c_TrackBlockHash).Width & "</TrackBlockHash_w>" & vbCrLf
            s_ConfigPAT_text &= "<Version_w>" & Me.SPS_P_ListView.Columns(c_Version).Width & "</Version_w>" & vbCrLf
            s_ConfigPAT_text &= "<LatestVersion_w>" & Me.SPS_P_ListView.Columns(c_LatestVersion).Width & "</LatestVersion_w>" & vbCrLf
            s_ConfigPAT_text &= "<ReleaseDate_w>" & Me.SPS_P_ListView.Columns(c_ReleaseDate).Width & "</ReleaseDate_w>" & vbCrLf
            s_ConfigPAT_text &= "<DownloadUrl_w>" & Me.SPS_P_ListView.Columns(c_DownloadUrl).Width & "</DownloadUrl_w>" & vbCrLf
            s_ConfigPAT_text &= "<DownloadSizeKb_w>" & Me.SPS_P_ListView.Columns(c_DownloadSizeKb).Width & "</DownloadSizeKb_w>" & vbCrLf
            s_ConfigPAT_text &= "<SPSCreationDate_w>" & Me.SPS_P_ListView.Columns(c_SPSCreationDate).Width & "</SPSCreationDate_w>" & vbCrLf
            s_ConfigPAT_text &= "<SPSModificationDate_w>" & Me.SPS_P_ListView.Columns(c_SPSModificationDate).Width & "</SPSModificationDate_w>" & vbCrLf
            s_ConfigPAT_text &= "<SPSPublisherName_w>" & Me.SPS_P_ListView.Columns(c_SPSPublisherName).Width & "</SPSPublisherName_w>" & vbCrLf
            s_ConfigPAT_text &= "<SuiteName_w>" & Me.SPS_P_ListView.Columns(c_SuiteName).Width & "</SuiteName_w>" & vbCrLf
            Dim colOrder As String = ""
            For i As Integer = 0 To SPS_P_ListView.Columns.Count - 1
                If i > 0 Then colOrder &= ","
                colOrder &= SPS_P_ListView.Columns(i).DisplayIndex.ToString()
            Next
            s_ConfigPAT_text &= "<ColumnOrder>" & colOrder & "</ColumnOrder>" & vbCrLf
        End If
        s_ConfigPAT_text &= "<DarkMode>" & DarkModeToolStripMenuItem.Checked.ToString() & "</DarkMode>" & vbCrLf
        s_ConfigPAT_text &= "<SplitOrientation>" & If(b_HorizontalLayout, "Horizontal", "Vertical") & "</SplitOrientation>" & vbCrLf        'Write the file
        File.WriteAllText(s_ConfigPAT_FilePath, s_ConfigPAT_text, Encoding.UTF8)
        ShowThemedMessageBox("Configuration saved!", "Save Config", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    
    Private Sub RestoreDefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreDefaultToolStripMenuItem.Click
        s_ConfigPAT_text = "<PAT_ListFile>" & Path.GetFileName(s_SPS_P_ListFile_Path) & "</PAT_ListFile>" & vbCrLf &
                "<Form1_x>" & "150" & "</Form1_x>" & vbCrLf &
                "<Form1_y>" & "150" & "</Form1_y>" & vbCrLf &
                "<Form1_w>" & "1100" & "</Form1_w>" & vbCrLf &
                "<Form1_h>" & "700" & "</Form1_h>" & vbCrLf &
                "<SplitterDistance>" & "750" & "</SplitterDistance>" & vbCrLf &
                "<SPS_Name_w>" & "225" & "</SPS_Name_w>" & vbCrLf &
                "<TrackURL_w>" & "130" & "</TrackURL_w>" & vbCrLf &
                "<TrackStartString_w>" & "150" & "</TrackStartString_w>" & vbCrLf &
                "<TrackStopString_w>" & "150" & "</TrackStopString_w>" & vbCrLf &
                "<TrackBlockHash_w>" & "150" & "</TrackBlockHash_w>" & vbCrLf &
                "<Version_w>" & "90" & "</Version_w>" & vbCrLf &
                "<LatestVersion_w>" & "90" & "</LatestVersion_w>" & vbCrLf &
                "<ReleaseDate_w>" & "120" & "</ReleaseDate_w>" & vbCrLf &
                "<DownloadUrl_w>" & "130" & "</DownloadUrl_w>" & vbCrLf &
                "<DownloadSizeKb_w>" & "85" & "</DownloadSizeKb_w>" & vbCrLf &
                "<SPSCreationDate_w>" & "120" & "</SPSCreationDate_w>" & vbCrLf &
                "<SPSModificationDate_w>" & "120" & "</SPSModificationDate_w>" & vbCrLf &
                "<SPSPublisherName_w>" & "120" & "</SPSPublisherName_w>" & vbCrLf &
                "<SuiteName_w>" & "120" & "</SuiteName_w>" & vbCrLf &
                "<ColumnOrder>0,1,2,3,4,5,6,7,8,9,10,11,12,13</ColumnOrder>" & vbCrLf
        REM Principal Form position and redimension
        Me.Width = S_GetsNBlockFromText(s_ConfigPAT_text, "<Form1_w>", "</Form1_w>", 1)
        Me.Height = S_GetsNBlockFromText(s_ConfigPAT_text, "<Form1_h>", "</Form1_h>", 1)
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) \ 2
        Dim i_SplitterDist As Integer = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SplitterDistance>", "</SplitterDistance>", 1))
        If i_SplitterDist > 100 Then
            SplitContainer1.SplitterDistance = i_SplitterDist
        End If
        REM SPS_P_ListView position and redimension
        Me.SPS_P_ListView.Columns(c_SPS_Name).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPS_Name_w>", "</SPS_Name_w>", 1))
        Me.SPS_P_ListView.Columns(c_TrackURL).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackURL_w>", "</TrackURL_w>", 1))
        Me.SPS_P_ListView.Columns(c_TrackStartString).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackStartString_w>", "</TrackStartString_w>", 1))
        Me.SPS_P_ListView.Columns(c_TrackStopString).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackStopString_w>", "</TrackStopString_w>", 1))
        Me.SPS_P_ListView.Columns(c_TrackBlockHash).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<TrackBlockHash_w>", "</TrackBlockHash_w>", 1))
        Me.SPS_P_ListView.Columns(c_Version).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<Version_w>", "</Version_w>", 1))
        Me.SPS_P_ListView.Columns(c_LatestVersion).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<LatestVersion_w>", "</LatestVersion_w>", 1))
        Me.SPS_P_ListView.Columns(c_ReleaseDate).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<ReleaseDate_w>", "</ReleaseDate_w>", 1))
        Me.SPS_P_ListView.Columns(c_DownloadUrl).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<DownloadUrl_w>", "</DownloadUrl_w>", 1))
        Me.SPS_P_ListView.Columns(c_DownloadSizeKb).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<DownloadSizeKb_w>", "</DownloadSizeKb_w>", 1))
        Me.SPS_P_ListView.Columns(c_SPSCreationDate).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSCreationDate_w>", "</SPSCreationDate_w>", 1))
        Me.SPS_P_ListView.Columns(c_SPSModificationDate).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSModificationDate_w>", "</SPSModificationDate_w>", 1))
        Me.SPS_P_ListView.Columns(c_SPSPublisherName).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SPSPublisherName_w>", "</SPSPublisherName_w>", 1))
        Me.SPS_P_ListView.Columns(c_SuiteName).Width = Convert.ToInt32(S_GetsNBlockFromText(s_ConfigPAT_text, "<SuiteName_w>", "</SuiteName_w>", 1))
        ' Reset column order to default
        For i As Integer = 0 To SPS_P_ListView.Columns.Count - 1
            SPS_P_ListView.Columns(i).DisplayIndex = i
        Next
        ' Reset View menu checkboxes to match default (all visible)
        ViewSPS_NameToolStripMenuItem.Checked = True
        ViewTrackURLToolStripMenuItem.Checked = True
        ViewTrackStartStringToolStripMenuItem.Checked = True
        ViewTrackStopStringToolStripMenuItem.Checked = True
        ViewTrackBlockHashToolStripMenuItem.Checked = True
        ViewVersionToolStripMenuItem.Checked = True
        ViewLatestVersionToolStripMenuItem.Checked = True
        ViewReleaseDateToolStripMenuItem.Checked = True
        ViewDownloadUrlToolStripMenuItem.Checked = True
        ViewDownloadSizeKbToolStripMenuItem.Checked = True
        ViewSPSCreationDateToolStripMenuItem.Checked = True
        ViewSPSModificationDateToolStripMenuItem.Checked = True
        ViewSPSPublisherNameToolStripMenuItem.Checked = True
        ViewSuiteNameToolStripMenuItem.Checked = True
        ' Reset orientation to vertical
        If b_HorizontalLayout Then
            b_HorizontalLayout = False
            ApplyLayoutOrientation()
        End If
    End Sub

    'ListView ToolStrip Subroutines
    Private Sub ToolStrip_OpenInBrowser_Click(sender As Object, e As EventArgs) Handles ToolStrip_OpenInBrowser.Click
        Dim item As ListViewItem
        For Each item In SPS_P_ListView.SelectedItems
            OpenInBrowser(item.Index)
        Next
    End Sub
    
    Private Sub OpenInBrowser(ByVal _SPS_order As Integer)
        Using BrowserProcess As New Process()
            With BrowserProcess.StartInfo
                .FileName = Me.SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackURL).Text
                .CreateNoWindow = True
                .UseShellExecute = True
                .WindowStyle = ProcessWindowStyle.Hidden
            End With
            BrowserProcess.Start()
        End Using
    End Sub
    
    Private Sub ToolStrip_OpenInSPSBuilder_Click(sender As Object, e As EventArgs) Handles ToolStrip_OpenInSPSBuilder.Click
        Dim indices As New List(Of Integer)
        For Each item As ListViewItem In SPS_P_ListView.SelectedItems
            indices.Add(item.Index)
        Next
        OpenMultipleInSPSBuilder(indices)
    End Sub

    Private Sub ToolStrip_OpenCheckedSPS_Click(sender As Object, e As EventArgs) Handles ToolStrip_OpenCheckedSPS.Click
        ' Adjust this to set the maximum number of SPS files to open at once:
        Dim maxOpen As Integer = 20

        If SPS_P_ListView.CheckedItems.Count = 0 Then
            ShowThemedMessageBox("No items are checked.", "Open checked SPS", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If SPS_P_ListView.CheckedItems.Count > maxOpen Then
            Dim result As DialogResult = ShowThemedMessageBox(
                "You have " & SPS_P_ListView.CheckedItems.Count & " items checked." & vbCrLf &
                "Only the first " & maxOpen & " will be opened to avoid performance issues." & vbCrLf & vbCrLf &
                "Continue?", "Open checked SPS", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.No Then Return
        End If

        Dim indices As New List(Of Integer)
        Dim count As Integer = 0
        For Each item As ListViewItem In SPS_P_ListView.CheckedItems
            If count >= maxOpen Then Exit For
            indices.Add(item.Index)
            count += 1
        Next
        OpenMultipleInSPSBuilder(indices)
    End Sub

    Private Sub OpenMultipleInSPSBuilder(ByVal indices As List(Of Integer))
        If indices.Count = 0 Then Return

        Dim spsBuilderExe As String = Path.Combine(s_SyMenuSuite_Path, "SPS_Builder_sps", "SPSBuilder.exe")
        If Not IO.File.Exists(spsBuilderExe) Then
            ShowThemedMessageBox("SPS Builder not found at:" & vbCrLf &
                            spsBuilderExe & vbCrLf & vbCrLf &
                            "SPS Builder must be installed in the SyMenuSuite folder.",
                            "SPS Builder not found",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim tmpFolder As String = Path.Combine(s_SyMenuSuite_Path, "_Trash", "_TmpPAT")
        If Not IO.Directory.Exists(tmpFolder) Then IO.Directory.CreateDirectory(tmpFolder)

        ' Build a list of what we need to find: (index, SPSName, suitePath, found file path)
        Dim needed As New List(Of Tuple(Of Integer, String, String, String))
        For Each idx As Integer In indices
            Dim s_SPSName As String = Me.SPS_P_ListView.Items(idx).SubItems(c_SPS_Name).Text
            Dim suitePath As String = CType(Me.SPS_P_ListView.Items(idx).Tag, String)
            needed.Add(Tuple.Create(idx, s_SPSName, suitePath, ""))
        Next

        ' Step 1: Build index of all existing files in temp folder
        Dim tmpFiles As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If Directory.Exists(tmpFolder) Then
            For Each f As String In Directory.GetFiles(tmpFolder, "*.sps")
                Dim name As String = Path.GetFileName(f)
                If Not tmpFiles.ContainsKey(name) Then
                    tmpFiles.Add(name, f)
                End If
            Next
        End If

        ' Step 2: Find each SPS by filename
        For i As Integer = 0 To needed.Count - 1
            Dim item = needed(i)
            Dim expectedFileName As String = item.Item2.Replace(" ", "_") & ".sps"

            ' First check temp folder index
            If tmpFiles.ContainsKey(expectedFileName) Then
                needed(i) = Tuple.Create(item.Item1, item.Item2, item.Item3, tmpFiles(expectedFileName))
                Continue For
            End If

            ' Then check loose .sps in _Cache folder
            Dim cachePath As String = Path.Combine(item.Item3, "_Cache")
            If Directory.Exists(cachePath) Then
                Dim cacheFile As String = Path.Combine(cachePath, expectedFileName)
                If IO.File.Exists(cacheFile) Then
                    Dim destFile As String = Path.Combine(tmpFolder, expectedFileName)
                    IO.File.Copy(cacheFile, destFile, True)
                    needed(i) = Tuple.Create(item.Item1, item.Item2, item.Item3, destFile)
                End If
            End If
        Next

        ' Step 3: Extract ALL zips from all needed _Cache folders (once per suite)
        Dim stillMissing As Boolean = needed.Any(Function(n) n.Item4 = "")
        Dim suitesChecked As New HashSet(Of String)
        If stillMissing Then
            Dim shellType As Type = Type.GetTypeFromProgID("Shell.Application", True)
            Dim shellObj As Object = Activator.CreateInstance(shellType)
            Dim zipCount As Integer = 0

            For Each item In needed
                If item.Item4 <> "" Then Continue For ' Already found
                Dim cachePath As String = Path.Combine(item.Item3, "_Cache")
                If suitesChecked.Contains(cachePath) Then Continue For
                suitesChecked.Add(cachePath)

                If Not Directory.Exists(cachePath) Then Continue For
                For Each s_fileName In Directory.GetFiles(cachePath, "*.zip")
                    Dim outputFolder As Object = shellType.InvokeMember("NameSpace", BindingFlags.InvokeMethod, Nothing, shellObj, New [Object]() {tmpFolder})
                    Dim inputZipFile As Object = shellType.InvokeMember("NameSpace", BindingFlags.InvokeMethod, Nothing, shellObj, New [Object]() {s_fileName})
                    outputFolder.CopyHere((inputZipFile.Items), 4)
                    zipCount += 1
                Next
            Next

            ' Wait for async CopyHere to finish
            If zipCount > 0 Then
                System.Threading.Thread.Sleep(1000)
                ' Wait until no new files appear for 2 seconds
                Dim lastCount As Integer = 0
                Dim stableTime As Integer = 0
                While stableTime < 2000
                    Dim currentCount As Integer = Directory.GetFiles(tmpFolder, "*.sps").Length
                    If currentCount = lastCount Then
                        stableTime += 500
                    Else
                        stableTime = 0
                        lastCount = currentCount
                    End If
                    System.Threading.Thread.Sleep(500)
                    Application.DoEvents()
                End While
            End If

            ' Now search extracted files for still-missing items
            For i As Integer = 0 To needed.Count - 1
                Dim item = needed(i)
                If item.Item4 <> "" Then Continue For
                For Each extractedFile In Directory.GetFiles(tmpFolder, "*.sps")
                    Try
                        Dim spsContent As String = File.ReadAllText(extractedFile, Encoding.UTF8)
                        Dim progName As String = S_GetsNBlockFromText(spsContent, "<ProgramName>", "</ProgramName>", 1)
                        If progName = item.Item2 Then
                            needed(i) = Tuple.Create(item.Item1, item.Item2, item.Item3, extractedFile)
                            Exit For
                        End If
                    Catch
                    End Try
                Next
            Next
        End If

        ' Step 4: Launch SPSBuilder for each found file (with delay between launches)
        Dim failed As New List(Of String)
        For Each item In needed
            If item.Item4 = "" OrElse Not IO.File.Exists(item.Item4) Then
                failed.Add(item.Item2)
                Continue For
            End If
            Try
                Dim p_SPSBuilder As New Process()
                p_SPSBuilder.StartInfo.FileName = spsBuilderExe
                p_SPSBuilder.StartInfo.WorkingDirectory = Path.GetDirectoryName(spsBuilderExe)
                p_SPSBuilder.StartInfo.Arguments = """" & item.Item4 & """"
                p_SPSBuilder.StartInfo.UseShellExecute = True
                p_SPSBuilder.StartInfo.WindowStyle = ProcessWindowStyle.Normal
                p_SPSBuilder.Start()
                p_SPSBuilder.WaitForInputIdle()
            Catch ex As Exception
                failed.Add(item.Item2 & " (" & ex.Message & ")")
            End Try
            ' Adjust time if it doesnt open all SPSBuilder windows:
            System.Threading.Thread.Sleep(70)
            Application.DoEvents()
        Next

        If failed.Count > 0 Then
            ShowThemedMessageBox("Could not open " & failed.Count & " SPS file(s):" & vbCrLf & vbCrLf &
                String.Join(vbCrLf, failed),
                "Open in SPS Builder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub ToolStrip_RefreshSelectedHash_Click(sender As Object, e As EventArgs) Handles ToolStrip_RefreshSelectedHash.Click
        Dim item As ListViewItem
        For Each item In SPS_P_ListView.SelectedItems
            ProgressBar1.Value = CInt(ProgressBar1.Minimum + (SPS_P_ListView.SelectedItems.IndexOf(item) + 1) * (ProgressBar1.Maximum - ProgressBar1.Minimum) / SPS_P_ListView.SelectedItems.Count)
            RefreshHash(item.Index)
        Next
    End Sub
    
    Private Sub RefreshHash(ByVal _SPS_order As Integer)
        REM Check SPS Downloading Track Block from the Web page
        Dim s_Web_Page As String
        Dim s_Web_Block As String
        Try
            Dim strOutput As String = ""
            Dim wrRequest As HttpWebRequest
            Dim wrResponse As HttpWebResponse
            Dim tempCookies As New CookieContainer
            wrRequest = CType(HttpWebRequest.Create(SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackURL).Text), HttpWebRequest)
            wrRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 Edg/128.0.0.0"
            wrRequest.Timeout = 10000
            wrRequest.CookieContainer = tempCookies
            wrResponse = CType(wrRequest.GetResponse(), HttpWebResponse)
            Using sr As New StreamReader(wrResponse.GetResponseStream())
                s_Web_Page = sr.ReadToEnd()
                sr.Close()
                wrResponse.Close()
            End Using
            REM Process web with Web Download successful
            s_Web_Block = S_GetsNBlockFromText(s_Web_Page, SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackStartString).Text, SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackStopString).Text, 1)
            If s_Web_Block = "" Then s_Web_Block = s_Web_Page
            REM Process block
            Using Md5 As New MD5CryptoServiceProvider()
                Dim ueCodigo As New UnicodeEncoding()
                Dim bHash() As Byte = Md5.ComputeHash(ueCodigo.GetBytes(s_Web_Block))
                SPS_P_ListView.Items(_SPS_order).UseItemStyleForSubItems = False
                SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackBlockHash).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
                SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackBlockHash).Text = Convert.ToBase64String(bHash)
                SPS_P_ListView.Refresh()
                SPS_P_ListFile_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
                SPS_P_ListFile_Name.Refresh()
            End Using
        Catch ex As Exception
            REM Web process error
            Beep()
            SPS_P_ListView.Items(_SPS_order).UseItemStyleForSubItems = False
            SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackURL).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
            SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackBlockHash).BackColor = If(DarkModeToolStripMenuItem.Checked, Color.Empty, Color.White)
            SPS_P_ListView.Items(_SPS_order).SubItems(c_TrackBlockHash).Text = ""
            SPS_P_ListView.Refresh()
        End Try
    End Sub
    
    Private Sub ToolStrip_DeleteSelectedTrack_Click(sender As Object, e As EventArgs) Handles ToolStrip_DeleteSelectedTrack.Click
        Dim result As Integer = ShowThemedMessageBox("Are you sure of delete the selected track lines?", "Delete Track", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Dim item As ListViewItem
            For Each item In SPS_P_ListView.SelectedItems
                DeleteTrack(item.Index)
            Next
        End If
    End Sub
    
    Private Sub DeleteTrack(ByVal _SPS_order As Integer)
        Me.SPS_P_ListView.Items.RemoveAt(Me.SPS_P_ListView.Items(_SPS_order).Index)
        Me.SPS_P_ListFile_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
        Me.SPS_P_ListFile_Name.Refresh()
    End Sub

    Private Sub ToolStrip_CheckSelected_Click(sender As Object, e As EventArgs) Handles ToolStrip_CheckSelected.Click
        For Each item As ListViewItem In SPS_P_ListView.SelectedItems
            item.Checked = True
        Next
    End Sub

    Private Sub ToolStrip_UncheckSelected_Click(sender As Object, e As EventArgs) Handles ToolStrip_UncheckSelected.Click
        For Each item As ListViewItem In SPS_P_ListView.SelectedItems
            item.Checked = False
        Next
    End Sub

    Private Sub ToolStrip_CheckAll_Click(sender As Object, e As EventArgs) Handles ToolStrip_CheckAll.Click
        For Each item As ListViewItem In SPS_P_ListView.Items
            item.Checked = True
        Next
    End Sub

    Private Sub ToolStrip_UncheckAll_Click(sender As Object, e As EventArgs) Handles ToolStrip_UncheckAll.Click
        For Each item As ListViewItem In SPS_P_ListView.Items
            item.Checked = False
        Next
    End Sub

    'General Subrutines and Functions
    Public Function S_GetsNBlockFromText(ByVal s_Text As String, ByVal s_Begin As String, ByVal s_End As String, ByVal i_Occurrence As Integer) As String
        REM Returns the block number i_Ocurrence occurrence from s_Text between s_Begin and s_End without delimiters
        Dim counter As Integer
        Dim i_BlockBeginPos As Integer = 1
        Dim i_BlockEndPos As Integer
        For counter = 1 To i_Occurrence Step 1
            i_BlockBeginPos = InStr(i_BlockBeginPos, s_Text, s_Begin)
            If i_BlockBeginPos = 0 Then
                S_GetsNBlockFromText = ""
                Exit Function
            Else
                i_BlockBeginPos = (i_BlockBeginPos + Len(s_Begin))
                i_BlockEndPos = InStr(i_BlockBeginPos, s_Text, s_End)
                If i_BlockEndPos = 0 Then
                    S_GetsNBlockFromText = ""
                    Exit Function
                ElseIf counter < i_Occurrence Then
                    i_BlockBeginPos = i_BlockEndPos + Len(s_End)
                End If
            End If
        Next
        S_GetsNBlockFromText = Mid(s_Text, i_BlockBeginPos, i_BlockEndPos - i_BlockBeginPos)
        Return S_GetsNBlockFromText
    End Function
    
    'View Menu - Header Hide/Show controls:
    Private Sub ViewSPS_NameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewSPS_NameToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 0 Then
            SPS_P_ListView.Columns(0).Width = If(ViewSPS_NameToolStripMenuItem.Checked, 225, 0)
        End If
    End Sub

    Private Sub ViewTrackURLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewTrackURLToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 1 Then
            SPS_P_ListView.Columns(1).Width = If(ViewTrackURLToolStripMenuItem.Checked, 130, 0)
        End If
    End Sub

    Private Sub ViewTrackStartStringToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewTrackStartStringToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 2 Then
            SPS_P_ListView.Columns(2).Width = If(ViewTrackStartStringToolStripMenuItem.Checked, 150, 0)
        End If
    End Sub

    Private Sub ViewTrackStopStringToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewTrackStopStringToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 3 Then
            SPS_P_ListView.Columns(3).Width = If(ViewTrackStopStringToolStripMenuItem.Checked, 150, 0)
        End If
    End Sub

    Private Sub ViewTrackBlockHashToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewTrackBlockHashToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 4 Then
            SPS_P_ListView.Columns(4).Width = If(ViewTrackBlockHashToolStripMenuItem.Checked, 150, 0)
        End If
    End Sub

    Private Sub ViewVersionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewVersionToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 5 Then
            SPS_P_ListView.Columns(5).Width = If(ViewVersionToolStripMenuItem.Checked, 90, 0)
        End If
    End Sub
    
    Private Sub ViewLatestVersionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewLatestVersionToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 6 Then
            SPS_P_ListView.Columns(6).Width = If(ViewLatestVersionToolStripMenuItem.Checked, 90, 0)
        End If
    End Sub

    Private Sub ViewReleaseDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewReleaseDateToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 7 Then
            SPS_P_ListView.Columns(7).Width = If(ViewReleaseDateToolStripMenuItem.Checked, 120, 0)
        End If
    End Sub

    Private Sub ViewDownloadUrlToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewDownloadUrlToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 8 Then
            SPS_P_ListView.Columns(8).Width = If(ViewDownloadUrlToolStripMenuItem.Checked, 130, 0)
        End If
    End Sub

    Private Sub ViewDownloadSizeKbToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewDownloadSizeKbToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 9 Then
            SPS_P_ListView.Columns(9).Width = If(ViewDownloadSizeKbToolStripMenuItem.Checked, 85, 0)
        End If
    End Sub

    Private Sub ViewSPSCreationDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewSPSCreationDateToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 10 Then
            SPS_P_ListView.Columns(10).Width = If(ViewSPSCreationDateToolStripMenuItem.Checked, 120, 0)
        End If
    End Sub

    Private Sub ViewSPSModificationDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewSPSModificationDateToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 11 Then
            SPS_P_ListView.Columns(11).Width = If(ViewSPSModificationDateToolStripMenuItem.Checked, 120, 0)
        End If
    End Sub

    Private Sub ViewSPSPublisherNameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewSPSPublisherNameToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 12 Then
            SPS_P_ListView.Columns(12).Width = If(ViewSPSPublisherNameToolStripMenuItem.Checked, 120, 0)
        End If
    End Sub
    
    Private Sub ViewSuiteNameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewSuiteNameToolStripMenuItem.Click
        If SPS_P_ListView.Columns.Count > 13 Then
            SPS_P_ListView.Columns(13).Width = If(ViewSuiteNameToolStripMenuItem.Checked, 120, 0)
        End If
    End Sub
    
    Private Sub bgWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        Dim checkedIndices As New List(Of Integer)
        Dim itemData As New List(Of Dictionary(Of String, String))
        
        ' Collect data from UI thread
        Me.Invoke(Sub()
            For Each item As ListViewItem In SPS_P_ListView.CheckedItems
                checkedIndices.Add(item.Index)
                Dim d As New Dictionary(Of String, String)
                d("TrackURL") = item.SubItems(c_TrackURL).Text
                d("StartString") = item.SubItems(c_TrackStartString).Text
                d("StopString") = item.SubItems(c_TrackStopString).Text
                d("BlockHash") = item.SubItems(c_TrackBlockHash).Text
                d("DownloadUrl") = item.SubItems(c_DownloadUrl).Text
                d("DownloadSizeKb") = item.SubItems(c_DownloadSizeKb).Text
                d("Version") = item.SubItems(c_Version).Text
                itemData.Add(d)
            Next
        End Sub)
        
        For i As Integer = 0 To checkedIndices.Count - 1
            ' Check if user requested cancellation
            If bgWorker.CancellationPending Then
                e.Cancel = True
                Exit For
            End If

            Dim idx As Integer = checkedIndices(i)
            Dim data As Dictionary(Of String, String) = itemData(i)
            
            ' This runs on background thread - no UI freezing
            Dim result As Dictionary(Of String, String) = CheckSPS_Download(
                data("TrackURL"), data("StartString"), data("StopString"), data("DownloadUrl"))
            
            ' Update UI with results
            Me.Invoke(Sub()
                If result("TrackSuccess") = "True" Then
                    ' Update latest version
                    If result("LatestVersion") <> "" Then
                        SPS_P_ListView.Items(idx).SubItems(c_LatestVersion).Text = result("LatestVersion")
                        If result("LatestVersion") = data("Version") Then
                            SPS_P_ListView.Items(idx).SubItems(c_LatestVersion).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
                        Else
                            SPS_P_ListView.Items(idx).SubItems(c_LatestVersion).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeOrange, m_LightModeOrange)
                        End If
                    Else
                        SPS_P_ListView.Items(idx).SubItems(c_LatestVersion).Text = ""
                        SPS_P_ListView.Items(idx).SubItems(c_LatestVersion).BackColor = Color.Empty
                    End If
                    
                    ' Update hash
                    SPS_P_ListView.Items(idx).SubItems(c_TrackBlockHash).Text = result("WebBlockHash")
                    If data("BlockHash") <> result("WebBlockHash") Then
                        Beep()
                        SPS_P_ListView.Items(idx).UseItemStyleForSubItems = False
                        SPS_P_ListView.Items(idx).SubItems(c_TrackURL).BackColor = If(DarkModeToolStripMenuItem.Checked, Color.Empty, Color.White)
                        SPS_P_ListView.Items(idx).SubItems(c_TrackBlockHash).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeOrange, m_LightModeOrange)
                        SPS_P_ListFile_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
                        SPS_P_ListFile_Name.Refresh()
                        If Open_Browser.Checked Then
                            Using BrowserProcess As New Process()
                                With BrowserProcess.StartInfo
                                    .FileName = data("TrackURL")
                                    .CreateNoWindow = True
                                    .UseShellExecute = True
                                    .WindowStyle = ProcessWindowStyle.Hidden
                                End With
                                BrowserProcess.Start()
                            End Using
                        End If
                    Else
                        SPS_P_ListView.Items(idx).UseItemStyleForSubItems = False
                        SPS_P_ListView.Items(idx).SubItems(c_TrackURL).BackColor = If(DarkModeToolStripMenuItem.Checked, Color.Empty, Color.White)
                        SPS_P_ListView.Items(idx).SubItems(c_TrackBlockHash).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
                    End If
                Else
                    Beep()
                    SPS_P_ListView.Items(idx).UseItemStyleForSubItems = False
                    SPS_P_ListView.Items(idx).SubItems(c_TrackURL).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
                End If
                
                ' Download size check
                If result("DownloadSuccess") = "True" Then
                    Dim remoteSizeKb As Long = CLng(result("RemoteSizeKb"))
                    Dim localSizeKb As Long = 0
                    Long.TryParse(data("DownloadSizeKb"), localSizeKb)
                    If remoteSizeKb = 0 Then
                        SPS_P_ListView.Items(idx).UseItemStyleForSubItems = False
                        SPS_P_ListView.Items(idx).SubItems(c_DownloadSizeKb).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
                    ElseIf localSizeKb > 0 AndAlso 0.1 > (Math.Abs(localSizeKb - remoteSizeKb) / remoteSizeKb) Then
                        SPS_P_ListView.Items(idx).UseItemStyleForSubItems = False
                        SPS_P_ListView.Items(idx).SubItems(c_DownloadUrl).BackColor = If(DarkModeToolStripMenuItem.Checked, Color.Empty, Color.White)
                        SPS_P_ListView.Items(idx).SubItems(c_DownloadSizeKb).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeGreen, m_LightModeGreen)
                    Else
                        Beep()
                        SPS_P_ListView.Items(idx).UseItemStyleForSubItems = False
                        SPS_P_ListView.Items(idx).SubItems(c_DownloadUrl).BackColor = If(DarkModeToolStripMenuItem.Checked, Color.Empty, Color.White)
                        SPS_P_ListView.Items(idx).SubItems(c_DownloadSizeKb).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeOrange, m_LightModeOrange)
                    End If
                Else
                    Beep()
                    SPS_P_ListView.Items(idx).UseItemStyleForSubItems = False
                    SPS_P_ListView.Items(idx).SubItems(c_DownloadUrl).BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
                End If
                
                SPS_P_ListView.Refresh()
            End Sub)
            
            bgWorker.ReportProgress(CInt((i + 1) * 100 / checkedIndices.Count))
        Next
    End Sub
    
    Private Sub bgWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs)
        ProgressBar1.Value = e.ProgressPercentage
        ProgressBar1.Refresh()
    End Sub

    Private Sub bgWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
        Track_Checked.Image = img_TrackCheckedOriginal
        Track_Checked.Text = ""
        Track_Checked.Enabled = True
        ReBuild_SPS_List.Enabled = True
        SelectAll_CheckBox.Visible = True
        LoadToolStripMenuItem.Enabled = True
        ProgressBar1.Value = 0
        SaveToolStripMenuItem.Enabled = True
        SaveAsToolStripMenuItem.Enabled = True
        SPS_P_ListView.Refresh()
        SPS_P_ListFile_Name.BackColor = If(DarkModeToolStripMenuItem.Checked, m_DarkModeRed, m_LightModeRed)
        File_Save.Enabled = True
        File_SaveAs.Enabled = True
        
        If e.Cancelled Then
            ShowThemedMessageBox("Track checking was stopped.", "Stopped")
        End If
    End Sub
    
    Private Sub DarkModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DarkModeToolStripMenuItem.Click
        If DarkModeToolStripMenuItem.Checked Then
            ApplyDarkMode()
        Else
            ApplyLightMode()
        End If
    End Sub

    Private Sub ApplyDarkMode()
        ' Main form
        Me.BackColor = m_DarkModeBackground30
        Me.ForeColor = Color.White
        
        ' MenuStrip
        MenuStrip1.BackColor = m_DarkModeBackground30
        MenuStrip1.ForeColor = Color.White
        For Each item As ToolStripMenuItem In MenuStrip1.Items
            ColorizeMenuItems(item, m_DarkModeBackground30, Color.White)
        Next
        
        ' ContextMenuStrips
        ContextMenuStrip1.BackColor = m_DarkModeBackground30
        ContextMenuStrip1.ForeColor = Color.White
        For Each item As ToolStripItem In ContextMenuStrip1.Items
            If TypeOf item Is ToolStripMenuItem Then
                Dim menuItem = CType(item, ToolStripMenuItem)
                menuItem.BackColor = m_DarkModeBackground30
                menuItem.ForeColor = Color.White
            ElseIf TypeOf item Is ToolStripSeparator Then
                item.BackColor = m_BorderColor
            End If
        Next

        ContextMenuStrip2.BackColor = m_DarkModeBackground30
        ContextMenuStrip2.ForeColor = Color.White
        For Each item As ToolStripItem In ContextMenuStrip2.Items
            If TypeOf item Is ToolStripMenuItem Then
                Dim menuItem = CType(item, ToolStripMenuItem)
                menuItem.BackColor = m_DarkModeBackground30
                menuItem.ForeColor = Color.White
            ElseIf TypeOf item Is ToolStripSeparator Then
                item.BackColor = m_BorderColor
            End If
        Next
        
        ' Re-apply colors to checked items
        For i = 0 To SPS_P_ListView.Items.Count - 1
            ' If TrackURL is white (light mode color), change to dark
            If SPS_P_ListView.Items(i).SubItems(c_TrackURL).BackColor = Color.White Then
                SPS_P_ListView.Items(i).SubItems(c_TrackURL).BackColor = m_DarkModeBackground45
            End If
            ' If DownloadUrl is white (light mode color), change to dark
            If SPS_P_ListView.Items(i).SubItems(c_DownloadUrl).BackColor = Color.White Then
                SPS_P_ListView.Items(i).SubItems(c_DownloadUrl).BackColor = m_DarkModeBackground45
            End If
        Next
        SPS_P_ListView.Refresh()

        ' Panels
        Panel_Left.BackColor = m_DarkModeBackground30
        Panel_Right.BackColor = m_DarkModeBackground30
        
        ' ListView
        SPS_P_ListView.BackColor = m_DarkModeBackground45
        SPS_P_ListView.ForeColor = Color.White
        
        ' RichTextBox
        RichTextBox1.BackColor = m_DarkModeBackground45
        RichTextBox1.ForeColor = m_RichText
        Toggle_HTMLView.BackColor = m_DarkModeBackground45
        Toggle_HTMLView.ForeColor = Color.White
        Toggle_HTMLView.FlatStyle = FlatStyle.Flat
        Toggle_HTMLView.FlatAppearance.BorderColor = m_BorderColor
        
        ' TextBoxes
        Track_URL.BackColor = m_DarkModeBackground45
        Track_URL.ForeColor = Color.White
        Start_String.BackColor = m_DarkModeBackground45
        Start_String.ForeColor = Color.White
        Stop_String.BackColor = m_DarkModeBackground45
        Stop_String.ForeColor = Color.White
        Find_String.BackColor = m_DarkModeBackground45
        Find_String.ForeColor = Color.White
        SPS_P_Tracker_Name.BackColor = m_DarkModeBackground45
        SPS_P_Tracker_Name.ForeColor = Color.White
        SPS_P_ListFile_Name.BackColor = m_DarkModeBackground45
        SPS_P_ListFile_Name.ForeColor = Color.White
        
        ' Labels
        Label_CurrentFile.BackColor = m_DarkModeBackground30
        Label_CurrentFile.ForeColor = Color.White
        Label_SearchBar.BackColor = m_DarkModeBackground30
        Label_SearchBar.ForeColor = Color.White
        Label_FindCount.BackColor = m_DarkModeBackground30
        Label_FindCount.ForeColor = m_RichText
        Label1.BackColor = m_DarkModeBackground30
        Label1.ForeColor = Color.White
        Label2.BackColor = m_DarkModeBackground30
        Label2.ForeColor = Color.White
        Label3.BackColor = m_DarkModeBackground30
        Label3.ForeColor = Color.White
        Label4.BackColor = m_DarkModeBackground30
        Label4.ForeColor = Color.White
        SPS_P_Publisher_Number.ForeColor = m_RichText

        ' All Buttons
        Open_File.BackColor = m_DarkModeBackground45
        Open_File.ForeColor = Color.White
        Open_File.FlatStyle = FlatStyle.Flat
        Open_File.FlatAppearance.BorderColor = m_BorderColor
        
        Track_Checked.BackColor = m_DarkModeBackground45
        Track_Checked.ForeColor = Color.White
        Track_Checked.FlatStyle = FlatStyle.Flat
        Track_Checked.FlatAppearance.BorderColor = m_BorderColor
        
        File_Save.BackColor = m_DarkModeBackground45
        File_Save.ForeColor = Color.White
        File_Save.FlatStyle = FlatStyle.Flat
        File_Save.FlatAppearance.BorderColor = m_BorderColor
        
        File_SaveAs.BackColor = m_DarkModeBackground45
        File_SaveAs.ForeColor = Color.White
        File_SaveAs.FlatStyle = FlatStyle.Flat
        File_SaveAs.FlatAppearance.BorderColor = m_BorderColor
        
        ReBuild_SPS_List.BackColor = m_DarkModeBackground45
        ReBuild_SPS_List.ForeColor = Color.White
        ReBuild_SPS_List.FlatStyle = FlatStyle.Flat
        ReBuild_SPS_List.FlatAppearance.BorderColor = m_BorderColor
        
        Toggle_RightPane.BackColor = m_DarkModeBackground45
        Toggle_RightPane.ForeColor = Color.White
        Toggle_RightPane.FlatStyle = FlatStyle.Flat
        Toggle_RightPane.FlatAppearance.BorderColor = m_BorderColor
        
        Toggle_SplitOrientation.BackColor = m_DarkModeBackground45
        Toggle_SplitOrientation.ForeColor = Color.White
        Toggle_SplitOrientation.FlatStyle = FlatStyle.Flat
        Toggle_SplitOrientation.FlatAppearance.BorderColor = m_BorderColor
        
        Help.BackColor = m_DarkModeBackground45
        Help.ForeColor = Color.White
        Help.FlatStyle = FlatStyle.Flat
        Help.FlatAppearance.BorderColor = m_BorderColor
        
        Browser_TrackURL.BackColor = m_DarkModeBackground45
        Browser_TrackURL.ForeColor = Color.White
        Browser_TrackURL.FlatStyle = FlatStyle.Flat
        Browser_TrackURL.FlatAppearance.BorderColor = m_BorderColor
        
        Download_Track_URL.BackColor = m_DarkModeBackground45
        Download_Track_URL.ForeColor = Color.White
        Download_Track_URL.FlatStyle = FlatStyle.Flat
        Download_Track_URL.FlatAppearance.BorderColor = m_BorderColor
        
        Save_Track.BackColor = m_DarkModeBackground45
        Save_Track.ForeColor = Color.White
        Save_Track.FlatStyle = FlatStyle.Flat
        Save_Track.FlatAppearance.BorderColor = m_BorderColor
        
        Go_To_Start_String.BackColor = m_DarkModeBackground45
        Go_To_Start_String.ForeColor = Color.White
        Go_To_Start_String.FlatStyle = FlatStyle.Flat
        Go_To_Start_String.FlatAppearance.BorderColor = m_BorderColor
        
        Go_To_Stop_String.BackColor = m_DarkModeBackground45
        Go_To_Stop_String.ForeColor = Color.White
        Go_To_Stop_String.FlatStyle = FlatStyle.Flat
        Go_To_Stop_String.FlatAppearance.BorderColor = m_BorderColor
        
        Search_From_Top.BackColor = m_DarkModeBackground45
        Search_From_Top.ForeColor = Color.White
        Search_From_Top.FlatStyle = FlatStyle.Flat
        Search_From_Top.FlatAppearance.BorderColor = m_BorderColor
        
        Search_From_CARET.BackColor = m_DarkModeBackground45
        Search_From_CARET.ForeColor = Color.White
        Search_From_CARET.FlatStyle = FlatStyle.Flat
        Search_From_CARET.FlatAppearance.BorderColor = m_BorderColor
        
        Reverse_From_CARET.BackColor = m_DarkModeBackground45
        Reverse_From_CARET.ForeColor = Color.White
        Reverse_From_CARET.FlatStyle = FlatStyle.Flat
        Reverse_From_CARET.FlatAppearance.BorderColor = m_BorderColor
        
        Reverse_From_BOTTOM.BackColor = m_DarkModeBackground45
        Reverse_From_BOTTOM.ForeColor = Color.White
        Reverse_From_BOTTOM.FlatStyle = FlatStyle.Flat
        Reverse_From_BOTTOM.FlatAppearance.BorderColor = m_BorderColor
        
        ' ProgressBar
        ProgressBar1.BackColor = m_DarkModeBackground45
        
        ' Checkboxes
        Open_Browser.BackColor = m_DarkModeBackground30
        Open_Browser.ForeColor = Color.White
        SelectAll_CheckBox.BackColor = m_DarkModeBackground30
        SelectAll_CheckBox.ForeColor = Color.White
    End Sub

    Private Sub ApplyLightMode()
        ' Main form
        Me.BackColor = Color.FromKnownColor(KnownColor.Control)
        Me.ForeColor = Color.Black
        
        ' MenuStrip
        MenuStrip1.BackColor = Color.FromKnownColor(KnownColor.Control)
        MenuStrip1.ForeColor = Color.Black
        For Each item As ToolStripMenuItem In MenuStrip1.Items
            ColorizeMenuItems(item, Color.FromKnownColor(KnownColor.Control), Color.Black)
        Next
        
        ' ContextMenuStrips
        ContextMenuStrip1.BackColor = Color.FromKnownColor(KnownColor.Control)
        ContextMenuStrip1.ForeColor = Color.Black
        For Each item As ToolStripItem In ContextMenuStrip1.Items
            If TypeOf item Is ToolStripMenuItem Then
                Dim menuItem = CType(item, ToolStripMenuItem)
                menuItem.BackColor = Color.FromKnownColor(KnownColor.Control)
                menuItem.ForeColor = Color.Black
            End If
        Next

        ContextMenuStrip2.BackColor = Color.FromKnownColor(KnownColor.Control)
        ContextMenuStrip2.ForeColor = Color.Black
        For Each item As ToolStripItem In ContextMenuStrip2.Items
            If TypeOf item Is ToolStripMenuItem Then
                Dim menuItem = CType(item, ToolStripMenuItem)
                menuItem.BackColor = Color.FromKnownColor(KnownColor.Control)
                menuItem.ForeColor = Color.Black
            End If
        Next
        
        ' Re-apply colors to checked items
        For i = 0 To SPS_P_ListView.Items.Count - 1
            ' If TrackURL is dark (dark mode color), change to white
            If SPS_P_ListView.Items(i).SubItems(c_TrackURL).BackColor = m_DarkModeBackground45 Then
                SPS_P_ListView.Items(i).SubItems(c_TrackURL).BackColor = Color.White
            End If
            ' If DownloadUrl is dark (dark mode color), change to white
            If SPS_P_ListView.Items(i).SubItems(c_DownloadUrl).BackColor = m_DarkModeBackground45 Then
                SPS_P_ListView.Items(i).SubItems(c_DownloadUrl).BackColor = Color.White
            End If
        Next
        SPS_P_ListView.Refresh()

        ' Panels
        Panel_Left.BackColor = Color.FromKnownColor(KnownColor.Control)
        Panel_Right.BackColor = Color.FromKnownColor(KnownColor.Control)
        
        ' ListView
        SPS_P_ListView.BackColor = Color.White
        SPS_P_ListView.ForeColor = Color.Black
        
        ' RichTextBox
        RichTextBox1.BackColor = Color.White
        RichTextBox1.ForeColor = Color.Blue
        Toggle_HTMLView.BackColor = Color.FromKnownColor(KnownColor.Control)
        Toggle_HTMLView.ForeColor = Color.Black
        Toggle_HTMLView.FlatStyle = FlatStyle.Standard
        
        ' TextBoxes
        Track_URL.BackColor = Color.White
        Track_URL.ForeColor = Color.Black
        Start_String.BackColor = Color.White
        Start_String.ForeColor = Color.Black
        Stop_String.BackColor = Color.White
        Stop_String.ForeColor = Color.Black
        Find_String.BackColor = Color.White
        Find_String.ForeColor = Color.Black
        SPS_P_Tracker_Name.BackColor = Color.White
        SPS_P_Tracker_Name.ForeColor = Color.Black
        SPS_P_ListFile_Name.BackColor = Color.FromKnownColor(KnownColor.Window)
        SPS_P_ListFile_Name.ForeColor = Color.Black

        ' Labels
        Label_CurrentFile.BackColor = Color.FromKnownColor(KnownColor.Control)
        Label_CurrentFile.ForeColor = Color.Black
        Label_SearchBar.BackColor = Color.FromKnownColor(KnownColor.Control)
        Label_SearchBar.ForeColor = Color.Black
        Label_FindCount.BackColor = Color.FromKnownColor(KnownColor.Control)
        Label_FindCount.ForeColor = Color.FromKnownColor(KnownColor.HotTrack)
        Label1.BackColor = Color.FromKnownColor(KnownColor.Control)
        Label1.ForeColor = Color.Black
        Label2.BackColor = Color.FromKnownColor(KnownColor.Control)
        Label2.ForeColor = Color.Black
        Label3.BackColor = Color.FromKnownColor(KnownColor.Control)
        Label3.ForeColor = Color.Black
        Label4.BackColor = Color.FromKnownColor(KnownColor.Control)
        Label4.ForeColor = Color.Black
        SPS_P_Publisher_Number.ForeColor = Color.FromKnownColor(KnownColor.HotTrack)
        
        ' All Buttons
        Open_File.BackColor = Color.FromKnownColor(KnownColor.Control)
        Open_File.ForeColor = Color.Black
        Open_File.FlatStyle = FlatStyle.Standard
        
        Track_Checked.BackColor = Color.FromKnownColor(KnownColor.Control)
        Track_Checked.ForeColor = Color.Black
        Track_Checked.FlatStyle = FlatStyle.Standard
        
        File_Save.BackColor = Color.FromKnownColor(KnownColor.Control)
        File_Save.ForeColor = Color.Black
        File_Save.FlatStyle = FlatStyle.Standard
        
        ReBuild_SPS_List.BackColor = Color.FromKnownColor(KnownColor.Control)
        ReBuild_SPS_List.ForeColor = Color.Black
        ReBuild_SPS_List.FlatStyle = FlatStyle.Standard
        
        Toggle_RightPane.BackColor = Color.FromKnownColor(KnownColor.Control)
        Toggle_RightPane.ForeColor = Color.Black
        Toggle_RightPane.FlatStyle = FlatStyle.Standard
        
        Toggle_SplitOrientation.BackColor = Color.FromKnownColor(KnownColor.Control)
        Toggle_SplitOrientation.ForeColor = Color.Black
        Toggle_SplitOrientation.FlatStyle = FlatStyle.Standard
        
        Help.BackColor = Color.FromKnownColor(KnownColor.Control)
        Help.ForeColor = Color.Black
        Help.FlatStyle = FlatStyle.Standard
        
        Browser_TrackURL.BackColor = Color.FromKnownColor(KnownColor.Control)
        Browser_TrackURL.ForeColor = Color.Black
        Browser_TrackURL.FlatStyle = FlatStyle.Standard
        
        Download_Track_URL.BackColor = Color.FromKnownColor(KnownColor.Control)
        Download_Track_URL.ForeColor = Color.Black
        Download_Track_URL.FlatStyle = FlatStyle.Standard
        
        Save_Track.BackColor = Color.FromKnownColor(KnownColor.Control)
        Save_Track.ForeColor = Color.Black
        Save_Track.FlatStyle = FlatStyle.Standard
        
        File_SaveAs.BackColor = Color.FromKnownColor(KnownColor.Control)
        File_SaveAs.ForeColor = Color.Black
        File_SaveAs.FlatStyle = FlatStyle.Standard
        
        Go_To_Start_String.BackColor = Color.FromKnownColor(KnownColor.Control)
        Go_To_Start_String.ForeColor = Color.Black
        Go_To_Start_String.FlatStyle = FlatStyle.Standard
        
        Go_To_Stop_String.BackColor = Color.FromKnownColor(KnownColor.Control)
        Go_To_Stop_String.ForeColor = Color.Black
        Go_To_Stop_String.FlatStyle = FlatStyle.Standard
        
        Search_From_Top.BackColor = Color.FromKnownColor(KnownColor.Control)
        Search_From_Top.ForeColor = Color.Black
        Search_From_Top.FlatStyle = FlatStyle.Standard
        
        Search_From_CARET.BackColor = Color.FromKnownColor(KnownColor.Control)
        Search_From_CARET.ForeColor = Color.Black
        Search_From_CARET.FlatStyle = FlatStyle.Standard
        
        Reverse_From_CARET.BackColor = Color.FromKnownColor(KnownColor.Control)
        Reverse_From_CARET.ForeColor = Color.Black
        Reverse_From_CARET.FlatStyle = FlatStyle.Standard
        
        Reverse_From_BOTTOM.BackColor = Color.FromKnownColor(KnownColor.Control)
        Reverse_From_BOTTOM.ForeColor = Color.Black
        Reverse_From_BOTTOM.FlatStyle = FlatStyle.Standard
        
        ' ProgressBar
        ProgressBar1.BackColor = Color.FromKnownColor(KnownColor.Control)
        
        ' Checkboxes
        Open_Browser.BackColor = Color.FromKnownColor(KnownColor.Control)
        Open_Browser.ForeColor = Color.Black
        SelectAll_CheckBox.BackColor = Color.FromKnownColor(KnownColor.Control)
        SelectAll_CheckBox.ForeColor = Color.Black
    End Sub
    
    Private Sub ColorizeMenuItems(item As ToolStripMenuItem, backColor As Color, foreColor As Color)
        item.BackColor = backColor
        item.ForeColor = foreColor
        For Each subItem As ToolStripItem In item.DropDownItems
            If TypeOf subItem Is ToolStripMenuItem Then
                Dim subMenu = CType(subItem, ToolStripMenuItem)
                subMenu.BackColor = backColor
                subMenu.ForeColor = foreColor
                ColorizeMenuItems(subMenu, backColor, foreColor)
            ElseIf TypeOf subItem Is ToolStripSeparator Then
                subItem.BackColor = m_BorderColor
            End If
        Next
    End Sub
End Class

REM Implements the manual sorting of items by columns. 
Class ListViewItemComparer
    Implements IComparer

    Private i_col As Integer
    Private bln_AscOrder As Boolean

    Public Sub New()
        i_col = 0
        bln_AscOrder = True
    End Sub

    Public Sub New(ByVal i_column As Integer, ByVal bln_Ascending As Boolean)
        i_col = i_column
        bln_AscOrder = bln_Ascending
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
    Implements IComparer.Compare
        Dim textX As String = CType(x, ListViewItem).SubItems(i_col).Text
        Dim textY As String = CType(y, ListViewItem).SubItems(i_col).Text
        Dim result As Integer

        ' Try numeric comparison first
        Dim numX As Double, numY As Double
        If Double.TryParse(textX, numX) AndAlso Double.TryParse(textY, numY) Then
            result = numX.CompareTo(numY)
        ElseIf DateTime.TryParse(textX, Nothing) AndAlso DateTime.TryParse(textY, Nothing) Then
            Dim dateX As DateTime = DateTime.Parse(textX)
            Dim dateY As DateTime = DateTime.Parse(textY)
            result = dateX.CompareTo(dateY)
        Else
            result = String.Compare(textX, textY)
        End If

        If Not bln_AscOrder Then
            result = -result
        End If

        Return result
    End Function
End Class

Public Class DoubleBufferedListView
    Inherits ListView
    
    Public Sub New()
        Me.DoubleBuffered = True
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
    End Sub
End Class

Public Class HelpForm
    Inherits Form

    Private rtb As RichTextBox

    Public Sub New(darkMode As Boolean, bgColor As Color)
        Me.Text = "Help - Quick Start Guide"
        Me.Width = 600
        Me.Height = 500
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        rtb = New RichTextBox()
        rtb.Dock = DockStyle.None
        rtb.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        rtb.Location = New Point(15, 15)
        rtb.Size = New Size(Me.ClientSize.Width - 30, Me.ClientSize.Height - 30)
        rtb.ReadOnly = True
        rtb.BorderStyle = BorderStyle.None
        rtb.ScrollBars = RichTextBoxScrollBars.Vertical

        If darkMode Then
            Me.BackColor = bgcolor
            rtb.BackColor = bgcolor
            rtb.ForeColor = Color.White
        Else
            Me.BackColor = bgcolor
            rtb.BackColor = bgcolor
            rtb.ForeColor = Color.Black
        End If

        rtb.Rtf = BuildHelpRtf(darkMode)
        Me.Controls.Add(rtb)
    End Sub

    Private Function BuildHelpRtf(darkMode As Boolean) As String
        Dim titleColor As String = If(darkMode, "\red100\green180\blue255", "\red0\green100\blue200")
        Dim headingColor As String = If(darkMode, "\red130\green210\blue130", "\red0\green130\blue0")
        Dim textColor As String = If(darkMode, "\red255\green255\blue255", "\red0\green0\blue0")
        Dim greenColor As String = If(darkMode, "\red100\green200\blue100", "\red0\green150\blue0")
        Dim redColor As String = If(darkMode, "\red255\green120\blue120", "\red200\green0\blue0")
        Dim orangeColor As String = If(darkMode, "\red255\green180\blue50", "\red220\green140\blue0")

        Dim rtf As String =
            "{\rtf1\ansi\deff0" &
            "{\fonttbl{\f0 Segoe UI;}}" &
            "{\colortbl ;" & titleColor & ";" & headingColor & ";" & textColor & ";" & greenColor & ";" & redColor & ";" & orangeColor & ";}" &
            "\viewkind4\uc1\f0\fs22" &
            "\cf1\b\fs28 SPS Published App Track x64\b0\fs22\par" &
            "\cf1\b\fs24 Quick Start Guide\b0\fs22\par" &
            "\par" &
            "\cf2\b Getting Started\b0\cf3\par" &
            "  \bullet  First, you need to set up your own \b AppSuite \b0 and \b SPS files \b0 if using PAT to update your own files.\par" &
            "  \bullet  Type the \b Publisher Name(s) \b0 in the top \b Search Bar \b0 for the software you want to track. Press Enter.\par" &
            "  \bullet  Select an item and type the web address into the \b Track URL \b0 field - this is the webpage where the version number or update information is published.\par" &
            "  \bullet  Click the \b Download Track URL \b0 button to obtain the page source.\par" &
            "  \bullet  Click the \b Save \b0 or \b Save As... \b0 buttons to save the PAT file at any time where changes have been detected.\par" &
            "\par" &
            "\cf2\b Finding Version Strings\b0\cf3\par" &
            "  \bullet  Open the \b Track URL \b0 in your browser using the \b Visit URL \b0 button. Locate the version number and type this into the \b Search \b0 field at the bottom right of the window.\par" &
            "  \bullet  Copy a unique text string that appears \b BEFORE\b0  the version number and paste it into the \b Start String \b0 field.\par" &
            "  \bullet  Copy a unique text string that appears \b AFTER\b0  the version number and paste it into the \b Stop String \b0 field.\par" &
            "  \bullet  \cf5\b IMPORTANT! \b0\cf3 Click the \b Save changes \b0 button \b BEFORE \b0 clicking a new item in the list or changes will be lost!.\par" &
            "\par" &
            "\cf2\b Checking for Updates\b0\cf3\par" &
            "  \bullet  Check (tick) the items you want to verify, then click \b Check for updates\b0. This will only get info for 'checked' items, unchecked items are ignored.\par" &
            "  \bullet  Results are colour coded:\par" &
            "      \cf4\b Green\b0\cf3  = no change detected (up to date).\par" &
            "      \cf5\b Red\b0\cf3  = change detected (possible update available). If other fields are Red, then its likely you need to update or save that info.\par" &
            "  \bullet  Save your tracking list using \b File > Save\b0  or \b File > Save As\b0 .\par" &
            "\par" &
            "\cf2\b Tips\b0\cf3\par" &
            "  \bullet  Use \b Check All\b0  / \b Uncheck All\b0  to quickly select or deselect all items.\par" &
            "  \bullet  Use \b Shift+Click \b0 to select items between first and second shift+click. Then Right click and select \b Check/Uncheck Selected\b0  to quickly select or deselect groups of items.\par" &
            "  \bullet  Use \b Ctrl+Click \b0 to select multiple items. Again, Right click and select \b Check/Uncheck Selected \b0 to quickly select or deselect groups of items.\par" &
            "  \bullet  Columns can be reordered by dragging the column headers.\par" &
            "  \bullet  Column visibility can be changed from the \b View \b0 menu.\par" &
            "  \bullet  Use \b File > Restore Defaults\b0  to reset window size, position, and column layout.\par" &
            "  \bullet  All window settings are saved on \b Exit \b0. Settings related to PAT files, requires manual save.\par" &
            "\par" &
            "\cf2\b Colour Guide\b0\cf3\par" &
            "\cf3 After checking for updates, columns are colour coded:\par" &
            "\par" &
            "\tx2000\tx4100\tx5800" &
            "\cf3\b Column\tab \cf5Red\tab \cf6Orange\tab \cf4Green\b0\par" &
            "  \cf3 Track URL\tab Connection error.\tab \tab Connected OK.\par" &
            "  \cf3 Track Block Hash\tab \tab Hash changed.\tab No change.\par" &
            "  \cf3 Latest Version\tab \tab New version.\tab Version matches.\par" &
            "  \cf3 Dwnld URL\tab Download failed.\tab \tab Download OK.\par" &
            "  \cf3 Dwnld KB\tab Size unknown (0).\tab Size changed.\tab Size matches.\par" &
            "\par" &
            "\cf3\b Summary:\b0\par" &
            "  \bullet  \cf5\b Red\b0\cf3  = Error (something failed or needs attention)\par" &
            "  \bullet  \cf6\b Orange\b0\cf3  = Change detected (update available)\par" &            "  \bullet  \cf4\b Green\b0\cf3  = OK (no change, everything matches)\par" &
            "\par" &
            "\cf1\b\fs24 Credits\b0\par" &
            "  \cf3 Original code by VVV_Easy_SyMenu.\par" &
            "  \cf3 Updated to v6.x.x.x 64bit by sl23.\par" &
            "}"

        Return rtf
    End Function
End Class