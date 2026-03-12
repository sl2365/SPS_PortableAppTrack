Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Xml

''' <summary>
''' Main application form.
''' Left pane  – ListView of SPS tracks loaded from a folder.
''' Right pane – Track details (editable) and a web-page viewer with
'''              text search / navigation.  Clicking a ListView item
'''              instantly populates the right pane.
''' </summary>
Public Class Form1

    '------------------------------------------------------------
    ' State
    '------------------------------------------------------------
    Private tracks As New List(Of SpsTrack)()
    Private currentTrack As SpsTrack = Nothing
    Private currentFolder As String = String.Empty

    ' Search state for the RichTextBox
    Private searchMatches As New List(Of Integer)()
    Private currentMatchIndex As Integer = -1

    ' Config file path (per-user AppData)
    Private ReadOnly configFilePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                     "SPS_PortableAppTrack", "config.xml")

    '============================================================
    ' Form Load / Close
    '============================================================

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Enable modern TLS for HTTPS downloads
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11
        LoadConfig()
        UpdateStatusBar()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SaveConfig()
    End Sub

    '============================================================
    ' Configuration – save / load window state, splitter positions,
    ' and the last-used folder.
    '============================================================

    Private Sub LoadConfig()
        Try
            If Not File.Exists(configFilePath) Then Return

            Dim doc As New XmlDocument()
            doc.Load(configFilePath)
            Dim root = doc.DocumentElement
            If root Is Nothing Then Return

            ' Window bounds
            Dim winNode = root.SelectSingleNode("Window")
            If winNode IsNot Nothing Then
                Dim left = XmlInt(winNode, "Left", Me.Left)
                Dim top = XmlInt(winNode, "Top", Me.Top)
                Dim width = XmlInt(winNode, "Width", Me.Width)
                Dim height = XmlInt(winNode, "Height", Me.Height)
                Dim maximized = XmlBool(winNode, "Maximized", False)

                Me.Left = left
                Me.Top = top
                Me.Width = Math.Max(900, width)
                Me.Height = Math.Max(500, height)
                If maximized Then Me.WindowState = FormWindowState.Maximized
            End If

            ' Splitter positions
            Dim splitNode = root.SelectSingleNode("Splitters")
            If splitNode IsNot Nothing Then
                splitMain.SplitterDistance = Math.Max(150, XmlInt(splitNode, "MainSplit", 380))
                splitRight.SplitterDistance = Math.Max(120, XmlInt(splitNode, "RightSplit", 230))
            End If

            ' Last folder
            Dim folderNode = root.SelectSingleNode("LastFolder")
            If folderNode IsNot Nothing AndAlso Directory.Exists(folderNode.InnerText) Then
                currentFolder = folderNode.InnerText
                LoadTracksFromFolder(currentFolder)
            End If

        Catch ex As Exception
            ' Config failures are non-critical – use defaults
        End Try
    End Sub

    Private Sub SaveConfig()
        Try
            Dim dir = Path.GetDirectoryName(configFilePath)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

            Dim doc As New XmlDocument()
            Dim root = doc.CreateElement("Config")
            doc.AppendChild(root)

            ' Window bounds (use RestoreBounds when maximised)
            Dim b = If(Me.WindowState = FormWindowState.Normal, Me.Bounds, Me.RestoreBounds)
            Dim winNode = AppendElement(doc, root, "Window")
            AppendElement(doc, winNode, "Left", b.Left.ToString())
            AppendElement(doc, winNode, "Top", b.Top.ToString())
            AppendElement(doc, winNode, "Width", b.Width.ToString())
            AppendElement(doc, winNode, "Height", b.Height.ToString())
            AppendElement(doc, winNode, "Maximized", (Me.WindowState = FormWindowState.Maximized).ToString())

            ' Splitter positions
            Dim splitNode = AppendElement(doc, root, "Splitters")
            AppendElement(doc, splitNode, "MainSplit", splitMain.SplitterDistance.ToString())
            AppendElement(doc, splitNode, "RightSplit", splitRight.SplitterDistance.ToString())

            ' Last folder
            AppendElement(doc, root, "LastFolder", currentFolder)

            Dim settings As New XmlWriterSettings() With {
                .Indent = True,
                .IndentChars = "  ",
                .Encoding = Encoding.UTF8
            }
            Using writer = XmlWriter.Create(configFilePath, settings)
                doc.Save(writer)
            End Using

        Catch ex As Exception
            ' Config save failure is non-critical
        End Try
    End Sub

    '------------------------------------------------------------
    ' XML helpers
    '------------------------------------------------------------

    Private Shared Function XmlInt(node As XmlNode, name As String, def As Integer) As Integer
        Dim child = node.SelectSingleNode(name)
        If child Is Nothing Then Return def
        Dim v As Integer
        Return If(Integer.TryParse(child.InnerText, v), v, def)
    End Function

    Private Shared Function XmlBool(node As XmlNode, name As String, def As Boolean) As Boolean
        Dim child = node.SelectSingleNode(name)
        If child Is Nothing Then Return def
        Dim v As Boolean
        Return If(Boolean.TryParse(child.InnerText, v), v, def)
    End Function

    Private Shared Function AppendElement(doc As XmlDocument, parent As XmlNode, name As String, Optional value As String = Nothing) As XmlElement
        Dim el = doc.CreateElement(name)
        If value IsNot Nothing Then el.InnerText = value
        parent.AppendChild(el)
        Return el
    End Function

    Private Shared Sub SetNodeText(doc As XmlDocument, parent As XmlNode, name As String, value As String)
        Dim node = parent.SelectSingleNode(name)
        If node Is Nothing Then
            node = doc.CreateElement(name)
            parent.AppendChild(node)
        End If
        node.InnerText = value
    End Sub

    Private Shared Function GetNodeText(parent As XmlNode, name As String) As String
        Dim node = parent.SelectSingleNode(name)
        Return If(node Is Nothing, String.Empty, node.InnerText.Trim())
    End Function

    '============================================================
    ' Open Folder
    '============================================================

    Private Sub OpenFolder()
        Using dlg As New FolderBrowserDialog()
            dlg.Description = "Select folder containing SPS files"
            If currentFolder <> String.Empty Then dlg.SelectedPath = currentFolder
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                currentFolder = dlg.SelectedPath
                LoadTracksFromFolder(currentFolder)
            End If
        End Using
    End Sub

    Private Sub LoadTracksFromFolder(folder As String)
        tracks.Clear()
        lvwTracks.Items.Clear()
        ClearRightPanel()

        Try
            Dim spsFiles = Directory.GetFiles(folder, "*.sps", SearchOption.TopDirectoryOnly)

            lvwTracks.BeginUpdate()
            For Each filePath In spsFiles
                Dim track = LoadSpsFile(filePath)
                If track IsNot Nothing Then
                    tracks.Add(track)
                    lvwTracks.Items.Add(BuildListViewItem(track))
                End If
            Next
            lvwTracks.EndUpdate()

            sslStatus.Text = $"Loaded {tracks.Count} track(s) from: {folder}"
            UpdateStatusBar()

        Catch ex As Exception
            MessageBox.Show($"Error loading folder:{Environment.NewLine}{ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sslStatus.Text = "Error loading folder"
        End Try
    End Sub

    '============================================================
    ' SPS File – Load / Save
    '============================================================

    Private Shared Function LoadSpsFile(filePath As String) As SpsTrack
        Try
            Dim doc As New XmlDocument()
            doc.Load(filePath)

            Dim track As New SpsTrack() With {.FilePath = filePath}

            ' Support both <SPS><Package>…</Package></SPS> and flat root
            Dim pkg = doc.SelectSingleNode("//Package")
            If pkg Is Nothing Then pkg = doc.DocumentElement
            If pkg Is Nothing Then Return Nothing

            track.PackageCode = GetNodeText(pkg, "PackageCode")
            track.Title = GetNodeText(pkg, "Title")
            track.Description = GetNodeText(pkg, "Description")
            track.Category = GetNodeText(pkg, "Category")
            track.SubCategory = GetNodeText(pkg, "SubCategory")
            track.Keywords = GetNodeText(pkg, "Keywords")
            track.ExternalVersion = GetNodeText(pkg, "ExternalVersion")
            track.ExternalVersionUrl = GetNodeText(pkg, "ExternalVersionUrl")
            track.HomePage = GetNodeText(pkg, "HomePage")
            track.InfoLink = GetNodeText(pkg, "InfoLink")
            track.DownloadLink = GetNodeText(pkg, "DownloadLink")
            track.Checksum = GetNodeText(pkg, "Checksum")

            ' Fall back to filename when Title is absent
            If String.IsNullOrWhiteSpace(track.Title) Then
                track.Title = Path.GetFileNameWithoutExtension(filePath)
            End If

            Return track

        Catch ex As Exception
            Return Nothing   ' Skip malformed files silently
        End Try
    End Function

    Private Sub SaveSpsFile(track As SpsTrack)
        Dim doc As New XmlDocument()

        ' Load the existing file to preserve any unknown fields
        If File.Exists(track.FilePath) Then
            doc.Load(track.FilePath)
        Else
            doc.AppendChild(doc.CreateXmlDeclaration("1.0", "utf-8", Nothing))
            Dim spsRoot = doc.CreateElement("SPS")
            doc.AppendChild(spsRoot)
            spsRoot.AppendChild(doc.CreateElement("Package"))
        End If

        Dim pkg = doc.SelectSingleNode("//Package")
        If pkg Is Nothing Then
            Dim spsRoot = doc.DocumentElement
            pkg = doc.CreateElement("Package")
            spsRoot.AppendChild(pkg)
        End If

        SetNodeText(doc, pkg, "Title", track.Title)
        SetNodeText(doc, pkg, "Description", track.Description)
        SetNodeText(doc, pkg, "ExternalVersion", track.ExternalVersion)
        SetNodeText(doc, pkg, "HomePage", track.HomePage)
        SetNodeText(doc, pkg, "DownloadLink", track.DownloadLink)

        Dim settings As New XmlWriterSettings() With {
            .Indent = True,
            .IndentChars = "  ",
            .Encoding = New UTF8Encoding(False)   ' UTF-8 without Byte Order Mark (BOM)
        }
        Using writer = XmlWriter.Create(track.FilePath, settings)
            doc.Save(writer)
        End Using
    End Sub

    '============================================================
    ' ListView helpers
    '============================================================

    Private Shared Function BuildListViewItem(track As SpsTrack) As ListViewItem
        Dim item As New ListViewItem(track.Title)
        item.SubItems.Add(track.ExternalVersion)
        item.SubItems.Add(track.Status)
        item.Tag = track
        Return item
    End Function

    '============================================================
    ' ListView – selection changed → populate right panel
    '============================================================

    Private Sub lvwTracks_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles lvwTracks.SelectedIndexChanged

        If lvwTracks.SelectedItems.Count = 0 Then
            currentTrack = Nothing
            Return
        End If

        currentTrack = TryCast(lvwTracks.SelectedItems(0).Tag, SpsTrack)
        If currentTrack IsNot Nothing Then PopulateRightPanel(currentTrack)
    End Sub

    Private Sub lvwTracks_DoubleClick(sender As Object, e As EventArgs) _
        Handles lvwTracks.DoubleClick

        ' Double-click jumps focus to the Name field for quick editing
        If currentTrack IsNot Nothing Then
            txtName.Focus()
            txtName.SelectAll()
        End If
    End Sub

    Private Sub PopulateRightPanel(track As SpsTrack)
        txtName.Text = track.Title
        txtVersion.Text = track.ExternalVersion
        txtHomePage.Text = track.HomePage
        txtDownloadLink.Text = track.DownloadLink
        txtDescription.Text = track.Description

        ' Pre-fill the URL box with the home page when it's empty
        ' or still pointing at the previous track
        If String.IsNullOrWhiteSpace(txtUrl.Text) Then
            txtUrl.Text = track.HomePage
        End If

        ' Clear the web pane since this is a different track
        rtbContent.Clear()
        ClearSearchState()

        sslStatus.Text = $"Selected: {track.Title}"
    End Sub

    Private Sub ClearRightPanel()
        txtName.Text = String.Empty
        txtVersion.Text = String.Empty
        txtHomePage.Text = String.Empty
        txtDownloadLink.Text = String.Empty
        txtDescription.Text = String.Empty
        txtUrl.Text = String.Empty
        rtbContent.Clear()
        ClearSearchState()
        currentTrack = Nothing
    End Sub

    '============================================================
    ' Download Web Page
    '============================================================

    Private Async Sub btnDownload_Click(sender As Object, e As EventArgs) _
        Handles btnDownload.Click

        Dim url = txtUrl.Text.Trim()
        If String.IsNullOrWhiteSpace(url) Then
            MessageBox.Show("Please enter a URL to download.",
                            "No URL", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim uri As Uri = Nothing
        If Not Uri.TryCreate(url, UriKind.Absolute, uri) OrElse
           (uri.Scheme <> Uri.UriSchemeHttp AndAlso uri.Scheme <> Uri.UriSchemeHttps) Then
            MessageBox.Show("Please enter a valid HTTP or HTTPS URL.",
                            "Invalid URL", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        btnDownload.Enabled = False
        btnDownload.Text = "Downloading…"
        sslStatus.Text = $"Downloading: {url}"
        rtbContent.Clear()
        ClearSearchState()

        Try
            Dim client As New WebClient()
            client.Encoding = Encoding.UTF8
            client.Headers.Add(HttpRequestHeader.UserAgent,
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " &
                "AppleWebKit/537.36 (KHTML, like Gecko) " &
                "Chrome/120.0 Safari/537.36")

            Dim html = Await client.DownloadStringTaskAsync(uri)
            Dim text = StripHtml(html)

            rtbContent.Text = text
            sslStatus.Text = $"Downloaded: {url}  ({text.Length:N0} characters)"

        Catch ex As WebException
            sslStatus.Text = $"Download failed: {ex.Message}"
            MessageBox.Show($"Could not download the page:{Environment.NewLine}{ex.Message}",
                            "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            sslStatus.Text = $"Error: {ex.Message}"
            MessageBox.Show($"Unexpected error:{Environment.NewLine}{ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnDownload.Enabled = True
            btnDownload.Text = "Download Page"
        End Try
    End Sub

    ''' <summary>Converts raw HTML to readable plain text.</summary>
    Private Shared Function StripHtml(html As String) As String
        ' Remove <script> and <style> blocks entirely
        Dim result = Regex.Replace(html, "<script[^>]*>[\s\S]*?</script>",
                                   String.Empty, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<style[^>]*>[\s\S]*?</style>",
                                String.Empty, RegexOptions.IgnoreCase)
        ' Convert block-level tags to newlines
        result = Regex.Replace(result, "<br\s*/?>", Environment.NewLine, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "</(p|div|h[1-6]|li|tr)>",
                                Environment.NewLine, RegexOptions.IgnoreCase)
        ' Strip remaining tags
        result = Regex.Replace(result, "<[^>]+>", String.Empty)
        ' Decode common HTML entities
        result = result.Replace("&amp;", "&")
        result = result.Replace("&lt;", "<")
        result = result.Replace("&gt;", ">")
        result = result.Replace("&quot;", """")
        result = result.Replace("&#39;", "'")
        result = result.Replace("&apos;", "'")
        result = result.Replace("&nbsp;", " ")
        ' Collapse excessive blank lines
        result = Regex.Replace(result, "(\r?\n){3,}", Environment.NewLine & Environment.NewLine)
        Return result.Trim()
    End Function

    Private Sub btnClearPage_Click(sender As Object, e As EventArgs) _
        Handles btnClearPage.Click

        rtbContent.Clear()
        ClearSearchState()
        sslStatus.Text = "Page cleared"
    End Sub

    '============================================================
    ' Text Search – find / navigate in RichTextBox
    '============================================================

    Private Sub ClearSearchState()
        searchMatches.Clear()
        currentMatchIndex = -1
        lblMatches.Text = "No matches"
        lblMatches.ForeColor = Drawing.Color.Gray
    End Sub

    ''' <summary>
    ''' Builds the full match list, highlights every occurrence in yellow,
    ''' and updates the match counter label.
    ''' </summary>
    Private Sub FindAllMatches(searchTerm As String)
        ClearSearchState()
        If String.IsNullOrEmpty(searchTerm) OrElse String.IsNullOrEmpty(rtbContent.Text) Then Return

        Dim textLower = rtbContent.Text.ToLowerInvariant()
        Dim termLower = searchTerm.ToLowerInvariant()

        Dim pos = 0
        Do
            pos = textLower.IndexOf(termLower, pos, StringComparison.Ordinal)
            If pos = -1 Then Exit Do
            searchMatches.Add(pos)
            pos += termLower.Length
        Loop

        If searchMatches.Count > 0 Then
            lblMatches.Text = $"{searchMatches.Count} match{If(searchMatches.Count = 1, "", "es")}"
            lblMatches.ForeColor = Drawing.Color.LimeGreen
            HighlightAllMatches(searchTerm)
        Else
            lblMatches.Text = "No matches"
            lblMatches.ForeColor = Drawing.Color.Red
        End If
    End Sub

    ''' <summary>Paints all matches yellow, leaving the rest in the default colours.</summary>
    Private Sub HighlightAllMatches(searchTerm As String)
        If searchMatches.Count = 0 Then Return

        ' Reset the entire text to default colours first
        rtbContent.SelectAll()
        rtbContent.SelectionBackColor = Drawing.Color.Black
        rtbContent.SelectionColor = Drawing.Color.Lime

        ' Highlight each match
        For Each pos In searchMatches
            rtbContent.Select(pos, searchTerm.Length)
            rtbContent.SelectionBackColor = Drawing.Color.Yellow
            rtbContent.SelectionColor = Drawing.Color.Black
        Next

        rtbContent.Select(0, 0)
    End Sub

    ''' <summary>
    ''' Scrolls to and highlights the match at <paramref name="index"/>
    ''' (wraps around).  Orange indicates the "current" match.
    ''' </summary>
    Private Sub NavigateToMatch(index As Integer)
        If searchMatches.Count = 0 Then Return

        ' Wrap
        If index < 0 Then index = searchMatches.Count - 1
        If index >= searchMatches.Count Then index = 0
        currentMatchIndex = index

        Dim pos = searchMatches(index)
        Dim termLen = txtFind.Text.Length

        ' Re-colour previous current match back to yellow
        HighlightAllMatches(txtFind.Text)

        ' Colour current match orange
        rtbContent.Select(pos, termLen)
        rtbContent.SelectionBackColor = Drawing.Color.Orange
        rtbContent.SelectionColor = Drawing.Color.Black
        rtbContent.ScrollToCaret()

        lblMatches.Text = $"{index + 1} of {searchMatches.Count}"
        lblMatches.ForeColor = Drawing.Color.LimeGreen
    End Sub

    Private Sub btnFindFirst_Click(sender As Object, e As EventArgs) _
        Handles btnFindFirst.Click

        FindAllMatches(txtFind.Text)
        If searchMatches.Count > 0 Then NavigateToMatch(0)
    End Sub

    Private Sub btnFindPrev_Click(sender As Object, e As EventArgs) _
        Handles btnFindPrev.Click

        If searchMatches.Count = 0 Then FindAllMatches(txtFind.Text)
        If searchMatches.Count > 0 Then NavigateToMatch(currentMatchIndex - 1)
    End Sub

    Private Sub btnFindNext_Click(sender As Object, e As EventArgs) _
        Handles btnFindNext.Click

        If searchMatches.Count = 0 Then FindAllMatches(txtFind.Text)
        If searchMatches.Count > 0 Then NavigateToMatch(currentMatchIndex + 1)
    End Sub

    Private Sub txtFind_KeyDown(sender As Object, e As KeyEventArgs) _
        Handles txtFind.KeyDown

        Select Case e.KeyCode
            Case Keys.Enter
                If e.Shift Then
                    btnFindPrev_Click(sender, e)
                Else
                    btnFindNext_Click(sender, e)
                End If
                e.Handled = True
            Case Keys.Escape
                ClearSearchState()
                ' Restore default colours
                rtbContent.SelectAll()
                rtbContent.SelectionBackColor = Drawing.Color.Black
                rtbContent.SelectionColor = Drawing.Color.Lime
                rtbContent.Select(0, 0)
                e.Handled = True
        End Select
    End Sub

    '============================================================
    ' Save Track
    '============================================================

    Private Sub SaveCurrentTrack()
        If currentTrack Is Nothing Then
            MessageBox.Show("Please select a track from the list first.",
                            "No Track Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Push UI values into the track object
        currentTrack.Title = txtName.Text.Trim()
        currentTrack.ExternalVersion = txtVersion.Text.Trim()
        currentTrack.HomePage = txtHomePage.Text.Trim()
        currentTrack.DownloadLink = txtDownloadLink.Text.Trim()
        currentTrack.Description = txtDescription.Text.Trim()

        Try
            SaveSpsFile(currentTrack)

            ' Refresh the ListView row
            Dim item = lvwTracks.SelectedItems(0)
            item.Text = currentTrack.Title
            item.SubItems(1).Text = currentTrack.ExternalVersion
            item.SubItems(2).Text = currentTrack.Status

            sslStatus.Text = $"Saved: {currentTrack.Title}"

        Catch ex As Exception
            MessageBox.Show($"Error saving track:{Environment.NewLine}{ex.Message}",
                            "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    '============================================================
    ' Add / Delete Track
    '============================================================

    Private Sub AddNewTrack()
        If String.IsNullOrWhiteSpace(currentFolder) Then
            MessageBox.Show("Please open a folder first before adding a track.",
                            "No Folder Open", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Deselect any current item
        lvwTracks.SelectedItems.Clear()
        ClearRightPanel()

        ' Build a new empty track
        Dim track As New SpsTrack() With {
            .FilePath = Path.Combine(currentFolder, "new_track.sps"),
            .Title = "New Track",
            .PackageCode = Guid.NewGuid().ToString(),
            .Status = "New"
        }
        currentTrack = track
        tracks.Add(track)

        Dim item = BuildListViewItem(track)
        lvwTracks.Items.Add(item)
        item.Selected = True
        item.EnsureVisible()

        ' Populate the right panel and place focus in the name field
        PopulateRightPanel(track)
        txtName.Focus()
        txtName.SelectAll()

        sslStatus.Text = "New track created – fill in details and click Save"
        UpdateStatusBar()
    End Sub

    Private Sub DeleteSelectedTrack()
        If lvwTracks.SelectedItems.Count = 0 Then
            MessageBox.Show("Please select a track to delete.",
                            "No Track Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim track = TryCast(lvwTracks.SelectedItems(0).Tag, SpsTrack)
        If track Is Nothing Then Return

        Dim answer = MessageBox.Show(
            $"Are you sure you want to delete '{track.Title}'?" &
            Environment.NewLine & "The SPS file will be permanently removed from disk.",
            "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If answer = DialogResult.Yes Then
            Try
                If File.Exists(track.FilePath) Then File.Delete(track.FilePath)
                tracks.Remove(track)
                lvwTracks.SelectedItems(0).Remove()
                ClearRightPanel()
                sslStatus.Text = $"Deleted: {track.Title}"
                UpdateStatusBar()
            Catch ex As Exception
                MessageBox.Show($"Error deleting track:{Environment.NewLine}{ex.Message}",
                                "Delete Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    '============================================================
    ' Open URL in the default browser
    '============================================================

    Private Sub OpenUrl(url As String)
        If String.IsNullOrWhiteSpace(url) Then Return
        Try
            Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show($"Could not open URL:{Environment.NewLine}{ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    '============================================================
    ' Status bar helper
    '============================================================

    Private Sub UpdateStatusBar()
        sslCount.Text = $"{tracks.Count} track{If(tracks.Count = 1, "", "s")}"
    End Sub

    '============================================================
    ' Toolbar handlers
    '============================================================

    Private Sub tsbOpenFolder_Click(sender As Object, e As EventArgs) _
        Handles tsbOpenFolder.Click
        OpenFolder()
    End Sub

    Private Sub tsbAddTrack_Click(sender As Object, e As EventArgs) _
        Handles tsbAddTrack.Click
        AddNewTrack()
    End Sub

    Private Sub tsbDeleteTrack_Click(sender As Object, e As EventArgs) _
        Handles tsbDeleteTrack.Click
        DeleteSelectedTrack()
    End Sub

    Private Sub tsbSave_Click(sender As Object, e As EventArgs) _
        Handles tsbSave.Click
        SaveCurrentTrack()
    End Sub

    '============================================================
    ' Menu handlers
    '============================================================

    Private Sub tsmiOpenFolder_Click(sender As Object, e As EventArgs) _
        Handles tsmiOpenFolder.Click
        OpenFolder()
    End Sub

    Private Sub tsmiSave_Click(sender As Object, e As EventArgs) _
        Handles tsmiSave.Click
        SaveCurrentTrack()
    End Sub

    Private Sub tsmiExit_Click(sender As Object, e As EventArgs) _
        Handles tsmiExit.Click
        Me.Close()
    End Sub

    Private Sub tsmiAddTrack_Click(sender As Object, e As EventArgs) _
        Handles tsmiAddTrack.Click
        AddNewTrack()
    End Sub

    Private Sub tsmiDeleteTrack_Click(sender As Object, e As EventArgs) _
        Handles tsmiDeleteTrack.Click
        DeleteSelectedTrack()
    End Sub

    Private Sub tsmiAbout_Click(sender As Object, e As EventArgs) _
        Handles tsmiAbout.Click
        MessageBox.Show(
            "SPS PortableAppTrack" & Environment.NewLine &
            "Version 1.0" & Environment.NewLine & Environment.NewLine &
            "Used alongside SyMenu / SPSBuilder to track portable" & Environment.NewLine &
            "app updates and manage SPS metadata files.",
            "About SPS PortableAppTrack",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)
    End Sub

    '============================================================
    ' Detail-pane button handlers
    '============================================================

    Private Sub btnSave_Click(sender As Object, e As EventArgs) _
        Handles btnSave.Click
        SaveCurrentTrack()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) _
        Handles btnClear.Click
        lvwTracks.SelectedItems.Clear()
        ClearRightPanel()
        sslStatus.Text = "Ready"
    End Sub

    Private Sub btnOpenHomePage_Click(sender As Object, e As EventArgs) _
        Handles btnOpenHomePage.Click
        OpenUrl(txtHomePage.Text.Trim())
    End Sub

    Private Sub txtHomePage_Leave(sender As Object, e As EventArgs) _
        Handles txtHomePage.Leave
        ' Auto-fill the URL box when it is empty
        If String.IsNullOrWhiteSpace(txtUrl.Text) Then
            txtUrl.Text = txtHomePage.Text.Trim()
        End If
    End Sub

    '============================================================
    ' Context menu – enable / disable items based on selection
    '============================================================

    Private Sub ctxListMenu_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) _
        Handles ctxListMenu.Opening

        Dim hasSelection = (lvwTracks.SelectedItems.Count > 0)
        ctxEdit.Enabled = hasSelection
        ctxOpenHome.Enabled = hasSelection
        ctxCopyHome.Enabled = hasSelection
        ctxCopyDownload.Enabled = hasSelection
        ctxDeleteTrack.Enabled = hasSelection
    End Sub

    Private Sub ctxEdit_Click(sender As Object, e As EventArgs) _
        Handles ctxEdit.Click
        If currentTrack IsNot Nothing Then
            txtName.Focus()
            txtName.SelectAll()
        End If
    End Sub

    Private Sub ctxOpenHome_Click(sender As Object, e As EventArgs) _
        Handles ctxOpenHome.Click
        If currentTrack IsNot Nothing Then OpenUrl(currentTrack.HomePage)
    End Sub

    Private Sub ctxCopyHome_Click(sender As Object, e As EventArgs) _
        Handles ctxCopyHome.Click
        If currentTrack IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(currentTrack.HomePage) Then
            Clipboard.SetText(currentTrack.HomePage)
            sslStatus.Text = "Home Page URL copied to clipboard"
        End If
    End Sub

    Private Sub ctxCopyDownload_Click(sender As Object, e As EventArgs) _
        Handles ctxCopyDownload.Click
        If currentTrack IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(currentTrack.DownloadLink) Then
            Clipboard.SetText(currentTrack.DownloadLink)
            sslStatus.Text = "Download URL copied to clipboard"
        End If
    End Sub

    Private Sub ctxDeleteTrack_Click(sender As Object, e As EventArgs) _
        Handles ctxDeleteTrack.Click
        DeleteSelectedTrack()
    End Sub

End Class

'================================================================
' SpsTrack – data model for one SPS metadata file
'================================================================

Public Class SpsTrack
    Public Property FilePath As String = String.Empty
    Public Property PackageCode As String = String.Empty
    Public Property Title As String = String.Empty
    Public Property Description As String = String.Empty
    Public Property Category As String = String.Empty
    Public Property SubCategory As String = String.Empty
    Public Property Keywords As String = String.Empty
    Public Property ExternalVersion As String = String.Empty
    Public Property ExternalVersionUrl As String = String.Empty
    Public Property HomePage As String = String.Empty
    Public Property InfoLink As String = String.Empty
    Public Property DownloadLink As String = String.Empty
    Public Property Checksum As String = String.Empty
    Public Property Status As String = String.Empty
    Public Property IsModified As Boolean = False
End Class
