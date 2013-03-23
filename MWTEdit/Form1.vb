Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D

Public Class Form1
    Private MstrVer = " 1.3"
    Private MobjXMLDoc As New MSXML2.DOMDocument
    Private MobjXMLDocFrag As New MSXML2.DOMDocument
    Private MobjXMLPages As MSXML2.IXMLDOMElement
    Friend MstrPage As String
    Private MobjXMLPage As MSXML2.IXMLDOMElement
    Private MstrHtmlTemplate As String = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN""><html><head><title></title><style type=""text/css"">body { background-color:black; color:#dcdcdc; font-family: Verdana; font-size:12pt; }</style></head><body>[TEXT]</body></html>"
    Private MobjButtons() As Button
    Private MblnDirty As Boolean
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
    Private MobjDomDoc As IHTMLDocument2
    Private MstrError As String
    Private MstrComment As String
    Private MobjLogWriter As System.IO.StreamWriter
    Private MblnDebug As Boolean = False
    Private thrMyThread As System.Threading.Thread = Nothing
    Private blnBackup As Boolean
    Private txtDirectory As String
    Private MstrImageFilter As String
    Private MstrAudioFilter As String
    Private MstrVideoFilter As String
    Private MstrFormTitle As String
    Private MintThumbnailSize As Integer
    Private MintLoopCheckDepth As Integer
    Private MintMaxDelay As Integer
    Private MstrEditStatus As String = "Neither"
    Dim WithEvents doc As System.Windows.Forms.HtmlDocument

    Private Sub loadFile(ByVal strFile As String)
        Try
            Dim objXMLEl As MSXML2.IXMLDOMElement
            TextBox1.Text = strFile
            MobjXMLDoc.load(TextBox1.Text)
            MobjXMLPages = MobjXMLDoc.selectSingleNode("//Pages")
            PopPageTree()
            tbMediaDirectory.Text = MobjXMLDoc.selectSingleNode("//MediaDirectory").text
            tbTitle.Text = MobjXMLDoc.selectSingleNode("//Title").text
            tbURL.Text = MobjXMLDoc.selectSingleNode("//Url").text
            tbMWTId.Text = getAttribute(MobjXMLDoc.documentElement, "id")
            objXMLEl = MobjXMLDoc.selectSingleNode("//Author")
            tbAuthorId.Text = getAttribute(objXMLEl, "id")
            tbAuthorName.Text = objXMLEl.selectSingleNode("Name").text
            tbAuthorURL.Text = objXMLEl.selectSingleNode("Url").text
            objXMLEl = MobjXMLDoc.selectSingleNode("//Settings")
            If Not objXMLEl Is Nothing Then
                cbAutoSetPageWhenSeen.Checked = objXMLEl.selectSingleNode("AutoSetPageWhenSeen").text
            Else
                cbAutoSetPageWhenSeen.Checked = False
            End If
            TreeViewPages.SelectedNode = TreeViewPages.Nodes(0)
            System.Windows.Forms.Application.DoEvents()
            fillListView()
            displaypage()
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", loadFile, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

    Private Sub fillListView()
        Try
            Cursor = Cursors.WaitCursor
            ListView1.Items.Clear()
            ListView2.Items.Clear()
            ListView3.Items.Clear()
            Dim myImageList = New ImageList()
            Dim strFiles() As String
            Dim strFile As String
            Dim intCount As Integer
            Dim img As Image
            Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
            Dim intProgress As Integer
            Dim objProgress As System.Windows.Forms.ToolStripProgressBar
            objProgress = StatusStrip1.Items("ToolStripProgressBar1")
            Dim intLoop As Integer

            myImageList.ImageSize = New Size(MintThumbnailSize, MintThumbnailSize)
            myImageList.ColorDepth = ColorDepth.Depth24Bit
            intCount = -1
            strFiles = GetPatternedFiles(MstrImageFilter.Split(","))
            intProgress = strFiles.Length
            ListView1.LargeImageList = myImageList
            For intLoop = 0 To strFiles.Length - 1
                Try
                    strFile = strFiles(intLoop)
                    If Not strFile Is Nothing Then
                        img = Image.FromFile(strFile)
                        StatusStrip1.Items("ToolStripStatusLabel1").Text = strFile
                        myImageList.Images.Add(GetPaddedAspectRatioThumbnail(img, myImageList.ImageSize))
                        img.Dispose()
                        ListView1.Items.Add(strFile)
                        objProgress.Value = intLoop / intProgress * 100
                        ListView1.Items(intLoop).ImageIndex = intLoop
                    End If
                Catch ex As Exception
                    MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", fillListView img, " & ex.Message & ", " & ex.TargetSite.Name)
                End Try
                System.Windows.Forms.Application.DoEvents()
            Next
            StatusStrip1.Items("ToolStripStatusLabel1").Text = ""
            objProgress.Value = 100

            strFiles = GetPatternedFiles(MstrAudioFilter.Split(","))
            For Each strFile In strFiles
                Try
                    ListView2.Items.Add(strFile)
                Catch ex As Exception
                End Try
                System.Windows.Forms.Application.DoEvents()
            Next

            strFiles = GetPatternedFiles(MstrVideoFilter.Split(","))
            For Each strFile In strFiles
                Try
                    ListView3.Items.Add(strFile)
                Catch ex As Exception
                End Try
                System.Windows.Forms.Application.DoEvents()
            Next

            Cursor = Cursors.Default
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", fillListView, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Function GetPaddedAspectRatioThumbnail(ByVal img As Image, ByVal newSize As Size) As Image

        Dim thumb As New Bitmap(newSize.Width, newSize.Height)
        Dim ratio As Double

        If img.Width > img.Height Then
            ratio = newSize.Width / img.Width
        Else
            ratio = newSize.Height / img.Height
        End If

        Using g As Graphics = Graphics.FromImage(thumb)
            'if you want to tweak the quality of the drawing...
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.SmoothingMode = SmoothingMode.HighQuality
            g.PixelOffsetMode = PixelOffsetMode.HighQuality
            g.CompositingQuality = CompositingQuality.HighQuality

            g.DrawImage(img, 0, 0, CInt(img.Width * ratio), CInt(img.Height * ratio))
        End Using

        Return thumb
    End Function

    Private Function GetPatternedFiles(ByVal strPatterns() As String) As String()
        Dim lrFiles As New ArrayList
        Dim strPattern As String
        Try
            For Each strPattern In strPatterns
                Dim strTemp() As String = System.IO.Directory.GetFiles(OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text, strPattern, IO.SearchOption.AllDirectories)
                If strTemp.Length > 0 Then
                    lrFiles.AddRange(strTemp)
                End If
            Next
            Dim strRet(lrFiles.Count) As String
            Array.Copy(lrFiles.ToArray, strRet, lrFiles.Count)
            Return strRet
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", GetPatternedFiles, " & ex.Message & ", " & ex.TargetSite.Name)
            GetPatternedFiles = Nothing
        End Try
    End Function

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MblnDebug Then
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", Form Closing")
        End If
        MobjLogWriter.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim strTemp As String
        MobjLogWriter = My.Computer.FileSystem.OpenTextFileWriter("ErrorLog.txt", True)
        If MblnDebug Then
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", Form Opening")
        End If
        Try
            MstrFormTitle = Me.Text & MstrVer
            Me.Text = MstrFormTitle
            OpenFileDialog1.DefaultExt = ".xml"
            OpenFileDialog1.Filter = "XML Files|*.xml"
            Dim strPath As String
            Dim objDomSettings As New MSXML2.DOMDocument
            Dim objXMLEl As MSXML2.IXMLDOMElement
            objDomSettings.load(My.Application.Info.DirectoryPath & "\Settings.xml")
            If objDomSettings.documentElement Is Nothing Then
                objXMLEl = objDomSettings.createElement("Settings")
                objXMLEl.setAttribute("directory", "c:\")
                objXMLEl.setAttribute("backup", "True")
                objXMLEl.setAttribute("thumbnailsize", "120")
                objXMLEl.setAttribute("loopcheckdepth", "20")
                objXMLEl.setAttribute("maxdelay", "120")
                objXMLEl.setAttribute("imagefilter", "*.jpg,*.bmp,*.gif")
                objXMLEl.setAttribute("audiofilter", "*.mp3,*.wav")
                objXMLEl.setAttribute("videofilter", "*.wmv,*.mp4")
                objDomSettings.documentElement = objXMLEl
                saveXml(objDomSettings, My.Application.Info.DirectoryPath & "\Settings.xml")
            End If
            strPath = getAttribute(objDomSettings.documentElement, "directory")
            blnBackup = getAttribute(objDomSettings.documentElement, "backup")
            txtDirectory = strPath
            If strPath <> "" Then
                OpenFileDialog1.InitialDirectory = strPath
            Else
                OpenFileDialog1.InitialDirectory = My.Application.Info.DirectoryPath
            End If

            strTemp = getAttribute(objDomSettings.documentElement, "thumbnailsize")
            If strTemp = "" Then
                MintThumbnailSize = 120
            Else
                MintThumbnailSize = Convert.ToInt32(strTemp)
            End If

            strTemp = getAttribute(objDomSettings.documentElement, "loopcheckdepth")
            If strTemp = "" Then
                MintLoopCheckDepth = 20
            Else
                MintLoopCheckDepth = Convert.ToInt32(strTemp)
            End If

            strTemp = getAttribute(objDomSettings.documentElement, "maxdelay")
            If strTemp = "" Then
                MintMaxDelay = 120
            Else
                MintMaxDelay = Convert.ToInt32(strTemp)
            End If

            MstrImageFilter = getAttribute(objDomSettings.documentElement, "imagefilter")
            If MstrImageFilter = "" Then
                MstrImageFilter = "*.jpg,*.bmp,*.gif"
            End If
            MstrAudioFilter = getAttribute(objDomSettings.documentElement, "audiofilter")
            If MstrAudioFilter = "" Then
                MstrAudioFilter = "*.mp3,*.wav"
            End If
            MstrVideoFilter = getAttribute(objDomSettings.documentElement, "videofilter")
            If MstrVideoFilter = "" Then
                MstrVideoFilter = "*.wmv,*.mp4"
            End If
            ReDim MobjButtons(0)
            MblnDirty = False
            AddHandler btnDelay.Click, AddressOf DynamicButtonClick
            WebBrowser3.DocumentText = "<html><head></head><body></body></html>"
            Do While Not WebBrowser3.ReadyState = WebBrowserReadyState.Complete
                System.Windows.Forms.Application.DoEvents()
            Loop
            txtRawText.Text = ""
            MobjDomDoc = WebBrowser3.Document.DomDocument
            MobjDomDoc.designMode = "On"
            AddHandler WebBrowser3.Document.ContextMenuShowing, New HtmlElementEventHandler(AddressOf Document_ContextMenuShowing)
            AddHandler tscbPaste.Click, AddressOf commandButton_Click
            AddHandler tscbCut.Click, AddressOf commandButton_Click
            AddHandler tscbCopy.Click, AddressOf commandButton_Click
            AddHandler tscbBold.Click, AddressOf commandButton_Click
            AddHandler tscbItalic.Click, AddressOf commandButton_Click
            AddHandler tscbUnderline.Click, AddressOf commandButton_Click
            DataGridView1.EnableHeadersVisualStyles = False
            DataGridView1.Columns(4).HeaderCell.Style.BackColor = Color.Red
            DataGridView1.Columns(5).HeaderCell.Style.BackColor = Color.Red
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", Form1_Load, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub PopPageTree()
        Try
            Dim objXMLPage As MSXML2.IXMLDOMNode
            Dim objXMLPageEl As MSXML2.IXMLDOMElement
            Dim objXMLError As MSXML2.IXMLDOMElement
            Dim objNode As System.Windows.Forms.TreeNode
            TabPage.Focus()
            System.Windows.Forms.Application.DoEvents()
            TreeViewPages.BeginUpdate()
            TreeViewPages.Nodes.Clear()
            For Each objXMLPage In MobjXMLPages.childNodes
                If objXMLPage.nodeType = MSXML2.DOMNodeType.NODE_ELEMENT Then
                    objXMLPageEl = objXMLPage
                    objNode = TreeViewPages.Nodes.Add(getAttribute(objXMLPageEl, "id"), getAttribute(objXMLPageEl, "id"))
                    objXMLError = objXMLPage.selectSingleNode("./Errors")
                    If Not objXMLError Is Nothing Then
                        objNode.ForeColor = Color.Red
                    End If
                End If
            Next
            TreeViewPages.EndUpdate()
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", PopPageTree, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub PopNextTree(ByRef objTreeNode As TreeNode, ByRef objXMLPage As MSXML2.IXMLDOMElement)
        Try
            Dim objXMLDelay As MSXML2.IXMLDOMElement
            Dim objXMLButtons As MSXML2.IXMLDOMNodeList
            Dim objXMLButton As MSXML2.IXMLDOMElement
            Dim objXMLPage2 As MSXML2.IXMLDOMElement
            Dim objTreeNode2 As TreeNode
            Dim objTreeNodeTest As TreeNode
            Dim strTarget As String
            Dim strButtonTarget As String
            objXMLDelay = objXMLPage.selectSingleNode("./Delay")
            If objXMLDelay Is Nothing Then
            Else
                strTarget = getAttribute(objXMLDelay, "target")
                objXMLPage2 = MobjXMLPages.selectSingleNode("Page[@id=""" & strTarget & """]")
                objTreeNode2 = objTreeNode.Nodes.Add(objXMLPage2.getAttribute("id"))
                System.Windows.Forms.Application.DoEvents()
                PopNextTree(objTreeNode2, objXMLPage2)
            End If

            objXMLButtons = objXMLPage.selectNodes("./Button")
            For intloop = 0 To objXMLButtons.length - 1
                objXMLButton = objXMLButtons.item(intloop)
                strButtonTarget = getAttribute(objXMLButton, "target")
                objTreeNodeTest = objTreeNode.Nodes.Item(strButtonTarget)
                If objTreeNodeTest Is Nothing Then
                    objXMLPage2 = MobjXMLPages.selectSingleNode("Page[@id=""" & strButtonTarget & """]")
                    objTreeNode2 = objTreeNode.Nodes.Add(getAttribute(objXMLPage2, "id"))
                    System.Windows.Forms.Application.DoEvents()
                    PopNextTree(objTreeNode2, objXMLPage2)
                End If
            Next
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", PopNextTree, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub TreeViewPages_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeViewPages.AfterSelect
        Try
            If MblnDirty Then
                Select Case MsgBox("Do you want to saved changes?" & vbCrLf & "Select Yes to save and move to the selected page, " & vbCrLf & "No lose changes or " & vbCrLf & "Cancel to to stay on this page", MsgBoxStyle.YesNoCancel, "Unsaved Changes")
                    Case MsgBoxResult.Yes
                        SavePage(False)
                        MblnDirty = False
                        displaypage()
                    Case MsgBoxResult.No
                        displaypage()
                End Select
            Else
                displaypage()
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", TreeViewPages_AfterSelect, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub bntImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImage.Click
        Try
            Dim strImage As String
            OpenFileDialogImage.InitialDirectory = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text
            If OpenFileDialogImage.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                If tbImage.Text <> OpenFileDialogImage.SafeFileName Then
                    tbImage.Text = OpenFileDialogImage.SafeFileName
                    strImage = OpenFileDialogImage.FileName
                    PictureBox1.Load(strImage)
                    PictureBox2.Load(strImage)
                    MblnDirty = True
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", bntImage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub cbMetronome_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMetronome.CheckedChanged
        MblnDirty = True
    End Sub

    Private Sub tbMetronome_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMetronome.TextChanged
        MblnDirty = True
    End Sub

    Private Sub cbDelay_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDelay.CheckedChanged
        MblnDirty = True
    End Sub

    Private Sub tbDelaySeconds_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelaySeconds.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbDelayTarget_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelayTarget.TextChanged
        MblnDirty = True
    End Sub

    Private Sub rbNormal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNormal.CheckedChanged
        MblnDirty = True
    End Sub

    Private Sub rbHidden_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbHidden.CheckedChanged
        MblnDirty = True
    End Sub

    Private Sub rbSecret_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSecret.CheckedChanged
        MblnDirty = True
    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        MblnDirty = True
    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        MblnDirty = True
    End Sub

    Private Sub DataGridView1_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        MblnDirty = True
    End Sub

    Private Sub txtPageSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPageSet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub txtPageUnSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPageUnSet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub txtPageIfSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPageIfSet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub txtPageIfNotSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPageIfNotSet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub txtDelaySet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDelaySet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub txtDelayUnSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDelayUnSet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub txtDelayIfSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDelayIfSet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub txtDelayIfNotSet_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDelayIfNotSet.TextChanged
        MblnDirty = True
    End Sub

    Private Sub btnAudio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAudio.Click
        Try
            OpenFileDialogImage.InitialDirectory = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text
            If OpenFileDialogImage.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                If tbAudio.Text <> OpenFileDialogImage.SafeFileName Then
                    tbAudio.Text = OpenFileDialogImage.SafeFileName
                    MblnDirty = True
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnAudio_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnVideo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVideo.Click
        Try
            OpenFileDialogImage.InitialDirectory = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text
            If OpenFileDialogImage.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                If tbVideo.Text <> OpenFileDialogImage.SafeFileName Then
                    tbVideo.Text = OpenFileDialogImage.SafeFileName
                    MblnDirty = True
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnVideo_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub cbAudio_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MblnDirty = True
    End Sub

    Private Sub cbVideo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MblnDirty = True
    End Sub

    Private Sub DynamicButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim objButton As Button
            Dim objNode As TreeNode
            objButton = sender
            For Each objNode In TreeViewPages.Nodes
                If objNode.Text = objButton.Tag Then
                    TreeViewPages.SelectedNode() = objNode
                    Exit For
                End If
            Next
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", DynamicButtonClick, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub SavePage()
        SavePage(True)
        MblnDirty = False
    End Sub

    Private Sub btnPlayAudio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlayAudio.Click
        Try
            Dim strAudio As String
            strAudio = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text & "\" & tbAudio.Text
            strAudio = Chr(34) & (strAudio) & Chr(34)
            mciSendString("open " & strAudio & " alias myDevice", Nothing, 0, 0)
            mciSendString("play myDevice", Nothing, 0, 0)
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnPlayAudio_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub savepage(ByVal blnRefresh As Boolean)
        Try
            'TODO save set / unset etc
            Dim objXMLImage As MSXML2.IXMLDOMElement
            Dim objXMLText As MSXML2.IXMLDOMElement
            Dim objXMLTextChild As MSXML2.IXMLDOMElement
            Dim objXMLDelay As MSXML2.IXMLDOMElement
            Dim objXMLMetronome As MSXML2.IXMLDOMElement
            Dim objXMLAudio As MSXML2.IXMLDOMElement
            Dim objXMLVideo As MSXML2.IXMLDOMElement
            Dim objXMLButtons As MSXML2.IXMLDOMNodeList
            Dim objXMLButton As MSXML2.IXMLDOMElement
            Dim strStyle As String
            Dim intloop As Integer

            'page set options
            addAttribute(MobjXMLPage, "set", txtPageSet.Text)
            addAttribute(MobjXMLPage, "unset", txtPageUnSet.Text)
            addAttribute(MobjXMLPage, "if-set", txtPageIfSet.Text)
            addAttribute(MobjXMLPage, "if-not-set", txtPageIfNotSet.Text)

            'Image
            objXMLImage = MobjXMLPage.selectSingleNode("Image")
            If Not objXMLImage Is Nothing Then
                objXMLImage.setAttribute("id", tbImage.Text)
            Else
                objXMLImage = MobjXMLDoc.createElement("Image")
                objXMLImage.setAttribute("id", tbImage.Text)
                MobjXMLPage.appendChild(objXMLImage)
            End If

            'Text
            objXMLText = MobjXMLPage.selectSingleNode("./Text")
            If objXMLText Is Nothing Then
                objXMLText = MobjXMLDoc.createElement("Text")
                MobjXMLPage.appendChild(objXMLText)
            End If
            strStyle = tbPageText.Text
            strStyle = strStyle.Replace("<BR>", "<BR/>")
            strStyle = strStyle.Replace("&nbsp;", "<BR/>")
            strStyle = strStyle.Trim(" ")
            strStyle = HtmlAsXml(strStyle)
            MobjXMLDocFrag.loadXML(strStyle)
            If MobjXMLDocFrag.documentElement Is Nothing Then
                strStyle = "<DIV>" & strStyle & "</DIV>"
                MobjXMLDocFrag.loadXML(strStyle)
            End If
            For intloop = objXMLText.childNodes.length - 1 To 0 Step -1
                If objXMLText.childNodes(intloop).nodeType = MSXML2.DOMNodeType.NODE_TEXT Then
                    objXMLText.text = ""
                Else
                    objXMLTextChild = objXMLText.childNodes(intloop)
                    objXMLText.removeChild(objXMLTextChild)
                End If
            Next
            objXMLText.appendChild(MobjXMLDocFrag.documentElement)

            'Delay
            objXMLDelay = MobjXMLPage.selectSingleNode("./Delay")
            If objXMLDelay Is Nothing Then
                If cbDelay.Checked Then
                    objXMLDelay = MobjXMLDoc.createElement("Delay")
                    MobjXMLPage.appendChild(objXMLDelay)
                    objXMLDelay.setAttribute("seconds", tbDelaySeconds.Text)
                    objXMLDelay.setAttribute("target", tbDelayTarget.Text)
                    objXMLDelay.setAttribute("start-with", tbDelayStartWith.Text)
                    Select Case True
                        Case rbHidden.Checked
                            strStyle = "hidden"
                        Case rbNormal.Checked
                            strStyle = "normal"
                        Case rbSecret.Checked
                            strStyle = "secret"
                        Case Else
                            strStyle = "normal"
                    End Select
                    objXMLDelay.setAttribute("style", strStyle)
                End If
            Else
                If cbDelay.Checked Then
                    objXMLDelay.setAttribute("seconds", tbDelaySeconds.Text)
                    objXMLDelay.setAttribute("target", tbDelayTarget.Text)
                    objXMLDelay.setAttribute("start-with", tbDelayStartWith.Text)
                    Select Case True
                        Case rbHidden.Checked
                            strStyle = "hidden"
                        Case rbNormal.Checked
                            strStyle = "normal"
                        Case rbSecret.Checked
                            strStyle = "secret"
                        Case Else
                            strStyle = "normal"
                    End Select
                    objXMLDelay.setAttribute("style", strStyle)
                    'delay set options
                    addAttribute(objXMLDelay, "set", txtDelaySet.Text)
                    addAttribute(objXMLDelay, "unset", txtDelayUnSet.Text)
                    addAttribute(objXMLDelay, "if-set", txtDelayIfSet.Text)
                    addAttribute(objXMLDelay, "if-not-set", txtDelayIfNotSet.Text)
                Else
                    MobjXMLPage.removeChild(objXMLDelay)
                End If
            End If

            'Metronome
            objXMLMetronome = MobjXMLPage.selectSingleNode("./Metronome")
            If objXMLMetronome Is Nothing Then
                If cbMetronome.Checked Then
                    objXMLMetronome = MobjXMLDoc.createElement("Metronome")
                    MobjXMLPage.appendChild(objXMLMetronome)
                    objXMLMetronome.setAttribute("bpm", tbMetronome.Text)
                End If
            Else
                If cbMetronome.Checked Then
                    objXMLMetronome.setAttribute("bpm", tbMetronome.Text)
                Else
                    MobjXMLPage.removeChild(objXMLMetronome)
                End If
            End If

            'Audio
            objXMLAudio = MobjXMLPage.selectSingleNode("./Audio")
            If objXMLAudio Is Nothing Then
                If tbAudio.Text <> "" Then
                    objXMLAudio = MobjXMLDoc.createElement("Audio")
                    MobjXMLPage.appendChild(objXMLAudio)
                    objXMLAudio.setAttribute("id", tbAudio.Text)
                    addAttribute(objXMLAudio, "target", tbAudioTarget.Text)
                    addAttribute(objXMLAudio, "start-at", tbAudioStartAt.Text)
                    addAttribute(objXMLAudio, "stop-at", tbAudioStopAt.Text)
                End If
            Else
                If tbAudio.Text <> "" Then
                    objXMLAudio.setAttribute("id", tbAudio.Text)
                    addAttribute(objXMLAudio, "target", tbAudioTarget.Text)
                    addAttribute(objXMLAudio, "start-at", tbAudioStartAt.Text)
                    addAttribute(objXMLAudio, "stop-at", tbAudioStopAt.Text)
                Else
                    MobjXMLPage.removeChild(objXMLAudio)
                End If
            End If

            'Video
            objXMLVideo = MobjXMLPage.selectSingleNode("./Video")
            If objXMLVideo Is Nothing Then
                If tbVideo.Text <> "" Then
                    objXMLVideo = MobjXMLDoc.createElement("Video")
                    MobjXMLPage.appendChild(objXMLVideo)
                    objXMLVideo.setAttribute("id", tbVideo.Text)
                    addAttribute(objXMLVideo, "target", tbVideoTarget.Text)
                    addAttribute(objXMLVideo, "start-at", tbVideoStartAt.Text)
                    addAttribute(objXMLVideo, "stop-at", tbVideoStopAt.Text)
                End If
            Else
                If tbVideo.Text <> "" Then
                    objXMLVideo.setAttribute("id", tbVideo.Text)
                    addAttribute(objXMLVideo, "target", tbVideoTarget.Text)
                    addAttribute(objXMLVideo, "start-at", tbVideoStartAt.Text)
                    addAttribute(objXMLVideo, "stop-at", tbVideoStopAt.Text)
                Else
                    MobjXMLPage.removeChild(objXMLVideo)
                End If
            End If

            'Buttons
            objXMLButtons = MobjXMLPage.selectNodes("./Button")
            For intloop = objXMLButtons.length - 1 To 0 Step -1
                objXMLButton = objXMLButtons.item(intloop)
                MobjXMLPage.removeChild(objXMLButton)
            Next
            For intloop = DataGridView1.Rows.Count - 2 To 0 Step -1
                objXMLButton = MobjXMLDoc.createElement("Button")
                objXMLButton.setAttribute("target", DataGridView1.Rows(intloop).Cells(1).Value)
                objXMLButton.text = DataGridView1.Rows(intloop).Cells(0).Value
                addAttribute(objXMLButton, "set", DataGridView1.Rows(intloop).Cells(2).Value)
                addAttribute(objXMLButton, "unset", DataGridView1.Rows(intloop).Cells(3).Value)
                addAttribute(objXMLButton, "if-set", DataGridView1.Rows(intloop).Cells(4).Value)
                addAttribute(objXMLButton, "if-not-set", DataGridView1.Rows(intloop).Cells(5).Value)
                MobjXMLPage.appendChild(objXMLButton)
            Next

            'Save XML File
            saveXml(MobjXMLDoc, TextBox1.Text)

            'Display Page
            If blnRefresh Then
                displaypage()
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", savepage, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Function HtmlAsXml(ByVal strHtml As String)
        Dim strXml As String = ""
        Try
            Dim intLoop As Integer
            Dim intPos As Integer
            Dim blnInTag As Boolean = False
            Dim blnInAtt As Boolean = False
            If Not strHtml Is Nothing Then
                For intLoop = 0 To strHtml.Length - 1
                    Select Case strHtml.Substring(intLoop, 1)
                        Case "<"
                            strXml = strXml & "<"
                            blnInTag = True
                        Case ">"
                            blnInTag = False
                            If blnInAtt Then
                                strXml = strXml & """>"
                                blnInAtt = False
                            Else
                                strXml = strXml & ">"
                            End If
                        Case "="
                            If blnInTag Then
                                If blnInAtt Then
                                    intPos = strXml.LastIndexOf(" ")
                                    strXml = strXml.Substring(0, intPos) & """ " & strXml.Substring(intPos)
                                End If
                                If strHtml.Substring(intLoop + 1, 1) = """" Then
                                    strXml = strXml & "="
                                    blnInAtt = False
                                Else
                                    strXml = strXml & "="""
                                    blnInAtt = True
                                End If
                            Else
                                strXml = strXml & "="
                            End If
                        Case Else
                            strXml = strXml & strHtml.Substring(intLoop, 1)
                    End Select
                Next
            Else
                strXml = "<DIV></DIV>"
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", HtmlAsXml, " & ex.Message & ", " & ex.TargetSite.Name)
        Finally
            HtmlAsXml = strXml
        End Try
    End Function

    Private Sub displaypage()
        Try
            Dim objXMLImage As MSXML2.IXMLDOMElement
            Dim objXMLText As MSXML2.IXMLDOMElement
            Dim objXMLDelay As MSXML2.IXMLDOMElement
            Dim objXMLMetronome As MSXML2.IXMLDOMElement
            Dim objXMLAudio As MSXML2.IXMLDOMElement
            Dim objXMLVideo As MSXML2.IXMLDOMElement
            Dim objXMLButtons As MSXML2.IXMLDOMNodeList
            Dim objXMLButton As MSXML2.IXMLDOMElement
            Dim objXMLError As MSXML2.IXMLDOMElement
            Dim objXMLComment As MSXML2.IXMLDOMNode
            Dim strImage As String
            Dim strText As String
            Dim strHtml As String
            Dim strSeconds As String
            Dim strTarget As String
            Dim strStyle As String
            Dim strBPM As String
            Dim strButtonTarget As String
            Dim strButtonText As String
            Dim strButtonSet As String
            Dim strButtonUnSet As String
            Dim strButtonIfSet As String
            Dim strButtonIfNotSet As String
            Dim intloop As Integer
            Dim intButtons As Integer
            Dim intSeconds As Integer
            Dim intMinutes As Integer
            Dim strFiles() As String
            Dim intRand As Integer
            Dim objRandom As New Random

            MstrPage = TreeViewPages.SelectedNode.Text
            Me.Text = MstrFormTitle & " " & MstrPage
            lblPage.Text = MstrPage
            lblPageName.Text = MstrPage
            MobjXMLPage = MobjXMLPages.selectSingleNode("./Page[@id=""" & MstrPage & """]")
            'page set options
            txtPageSet.Text = getAttribute(MobjXMLPage, "set")
            txtPageUnSet.Text = getAttribute(MobjXMLPage, "unset")
            txtPageIfSet.Text = getAttribute(MobjXMLPage, "if-set")
            txtPageIfNotSet.Text = getAttribute(MobjXMLPage, "if-not-set")
            'txtRawText.Text = ""

            'image
            objXMLImage = MobjXMLPage.selectSingleNode("Image")
            If Not objXMLImage Is Nothing Then
                tbImage.Text = getAttribute(objXMLImage, "id")
                If tbImage.Text = "" Then
                    If Not PictureBox1.Image Is Nothing Then
                        PictureBox1.Image.Dispose()
                        PictureBox1.Image = Nothing
                    End If
                    If Not PictureBox2.Image Is Nothing Then
                        PictureBox2.Image.Dispose()
                        PictureBox2.Image = Nothing
                    End If
                Else
                    If tbImage.Text.IndexOf("*") > -1 Then
                        strFiles = System.IO.Directory.GetFiles(OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text & "\", getAttribute(objXMLImage, "id"))
                        intRand = objRandom.Next(0, strFiles.Length)
                        strImage = strFiles(intRand)
                    Else
                        strImage = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text & "\" & getAttribute(objXMLImage, "id")
                    End If
                    PictureBox1.Load(strImage)
                    PictureBox2.Load(strImage)
                End If
            Else
                tbImage.Text = ""
                If Not PictureBox1.Image Is Nothing Then
                    PictureBox1.Image.Dispose()
                    PictureBox1.Image = Nothing
                End If
                If Not PictureBox2.Image Is Nothing Then
                    PictureBox2.Image.Dispose()
                    PictureBox2.Image = Nothing
                End If
            End If

            'text
            objXMLText = MobjXMLPage.selectSingleNode("./Text")
            If objXMLText Is Nothing Then
                strText = ""
                strHtml = MstrHtmlTemplate.Replace("[TEXT]", strText)
            Else
                strText = objXMLText.xml
                strText = strText.Replace("<Text>", "")
                strText = strText.Replace("</Text>", "")
                strHtml = MstrHtmlTemplate.Replace("[TEXT]", strText)
            End If
            WebBrowser1.DocumentText = strHtml
            Do While Not WebBrowser1.ReadyState = WebBrowserReadyState.Complete
                System.Windows.Forms.Application.DoEvents()
            Loop
            If Not WebBrowser1.Document.Body Is Nothing Then
                tbPageText.Text = WebBrowser1.Document.Body.InnerHtml
            End If
            'If Not WebBrowser1.Document.Body Is Nothing Then
            '    txtRawText.Text = WebBrowser1.Document.Body.InnerHtml
            'End If

            'Delay
            objXMLDelay = MobjXMLPage.selectSingleNode("./Delay")
            If objXMLDelay Is Nothing Then
                cbDelay.Checked = False
                tbDelaySeconds.Text = ""
                tbDelayTarget.Text = ""
                tbDelayStartWith.Text = ""
                rbHidden.Checked = False
                rbNormal.Checked = False
                rbSecret.Checked = False
                btnDelay.Tag = ""
                btnDelay.Enabled = False
                lblTimer.Text = ""
                txtDelaySet.Text = ""
                txtDelayUnSet.Text = ""
                txtDelayIfSet.Text = ""
                txtDelayIfNotSet.Text = ""
            Else
                cbDelay.Checked = True
                strSeconds = getAttribute(objXMLDelay, "seconds")
                strTarget = getAttribute(objXMLDelay, "target")
                strStyle = getAttribute(objXMLDelay, "style")
                tbDelaySeconds.Text = strSeconds
                tbDelayTarget.Text = strTarget
                tbDelayStartWith.Text = getAttribute(objXMLDelay, "start-with")
                btnDelay.Tag = strTarget
                btnDelay.Enabled = True
                If strSeconds.IndexOf("..") > -1 Then
                    lblTimer.Text = strSeconds & " " & strStyle
                Else
                    intSeconds = strSeconds
                    intMinutes = Math.Floor(intSeconds / 60)
                    intSeconds = intSeconds - (intMinutes * 60)
                    lblTimer.Text = Microsoft.VisualBasic.Right("0" & intMinutes, 2) & ":" & Microsoft.VisualBasic.Right("0" & intSeconds, 2) & " " & strStyle
                End If
                Select Case strStyle
                    Case "normal"
                        rbHidden.Checked = False
                        rbNormal.Checked = True
                        rbSecret.Checked = False
                    Case "hidden"
                        rbHidden.Checked = True
                        rbNormal.Checked = False
                        rbSecret.Checked = False
                    Case "secret"
                        rbHidden.Checked = False
                        rbNormal.Checked = False
                        rbSecret.Checked = True
                End Select
                'delay set options
                txtDelaySet.Text = getAttribute(objXMLDelay, "set")
                txtDelayUnSet.Text = getAttribute(objXMLDelay, "unset")
                txtDelayIfSet.Text = getAttribute(objXMLDelay, "if-set")
                txtDelayIfNotSet.Text = getAttribute(objXMLDelay, "if-not-set")
            End If

            'Metronome
            objXMLMetronome = MobjXMLPage.selectSingleNode("./Metronome")
            If objXMLMetronome Is Nothing Then
                cbMetronome.Checked = False
                tbMetronome.Text = ""
            Else
                cbMetronome.Checked = True
                strBPM = getAttribute(objXMLMetronome, "bpm")
                tbMetronome.Text = strBPM
            End If

            'Audio
            objXMLAudio = MobjXMLPage.selectSingleNode("./Audio")
            If objXMLAudio Is Nothing Then
                tbAudio.Text = ""
                tbAudioTarget.Text = ""
                tbAudioStartAt.Text = ""
                tbAudioStopAt.Text = ""
                btnPlayAudio.Enabled = False
            Else
                tbAudio.Text = getAttribute(objXMLAudio, "id")
                tbAudioTarget.Text = getAttribute(objXMLAudio, "target")
                tbAudioStartAt.Text = getAttribute(objXMLAudio, "start-at")
                tbAudioStopAt.Text = getAttribute(objXMLAudio, "stop-at")
                btnPlayAudio.Enabled = True
            End If

            'Video
            objXMLVideo = MobjXMLPage.selectSingleNode("./Video")
            If objXMLVideo Is Nothing Then
                tbVideo.Text = ""
                tbVideoTarget.Text = ""
                tbVideoStartAt.Text = ""
                tbVideoStopAt.Text = ""
            Else
                tbVideo.Text = getAttribute(objXMLVideo, "id")
                tbVideoTarget.Text = getAttribute(objXMLVideo, "target")
                tbVideoStartAt.Text = getAttribute(objXMLVideo, "start-at")
                tbVideoStopAt.Text = getAttribute(objXMLVideo, "stop-at")
            End If

            'Buttons
            objXMLButtons = MobjXMLPage.selectNodes("./Button")
            'clear buttons from previous page
            DataGridView1.Rows.Clear()
            For intloop = MobjButtons.GetUpperBound(0) To 1 Step -1
                FlowLayoutPanel1.Controls.Remove(MobjButtons(intloop))
                MobjButtons(intloop).Dispose()
            Next
            'populate with buttons for this page
            For intloop = objXMLButtons.length - 1 To 0 Step -1
                objXMLButton = objXMLButtons.item(intloop)
                strButtonTarget = getAttribute(objXMLButton, "target")
                strButtonSet = getAttribute(objXMLButton, "set")
                strButtonUnSet = getAttribute(objXMLButton, "unset")
                strButtonIfSet = getAttribute(objXMLButton, "if-set")
                strButtonIfNotSet = getAttribute(objXMLButton, "if-not-set")
                strButtonText = objXMLButton.text
                DataGridView1.Rows.Add(strButtonText, strButtonTarget, strButtonSet, strButtonUnSet, strButtonIfSet, strButtonIfNotSet)
            Next
            ReDim MobjButtons(0)
            intButtons = 0
            For intloop = 0 To objXMLButtons.length - 1
                objXMLButton = objXMLButtons.item(intloop)
                strButtonTarget = getAttribute(objXMLButton, "target")
                strButtonText = objXMLButton.text
                intButtons = intButtons + 1
                ReDim Preserve MobjButtons(intButtons)
                MobjButtons(intButtons) = New Button
                MobjButtons(intButtons).Text = strButtonText
                MobjButtons(intButtons).AutoSize = True
                MobjButtons(intButtons).Tag = strButtonTarget
                AddHandler MobjButtons(intButtons).Click, AddressOf DynamicButtonClick
                FlowLayoutPanel1.Controls.Add(MobjButtons(intButtons))
            Next

            'Error Nodes from downloaded tease
            objXMLError = MobjXMLPage.selectSingleNode("./Errors")
            If objXMLError Is Nothing Then
                MstrComment = ""
                MstrError = ""
                MenuPageDownLoadError.Enabled = False
                MenuPageDeleteError.Enabled = False
            Else
                MstrError = objXMLError.xml
                MstrComment = ""
                For Each objXMLComment In MobjXMLPage.childNodes
                    If objXMLComment.nodeType = MSXML2.DOMNodeType.NODE_COMMENT Then
                        MstrComment = MstrComment & objXMLComment.xml
                    End If
                Next
                MenuPageDownLoadError.Enabled = True
                MenuPageDeleteError.Enabled = True
            End If

            MblnDirty = False
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", displaypage, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub Document_ContextMenuShowing(ByVal sender As Object, ByVal e As HtmlElementEventArgs)

    End Sub

    Private Sub commandButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim objButton As ToolStripButton
            objButton = sender
            MobjDomDoc.execCommand(objButton.Tag.ToString(), False, Nothing)
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", commandButton_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub tsbtnColour_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbtnColour.Click
        Try
            Dim intColour As Integer
            Dim strColour As String
            Dim strSize As String
            Dim intSize As Integer
            Dim strFamily As String
            Dim objFont As Drawing.Font
            Dim objFontStyle As System.Drawing.FontStyle
            Dim blnBold As Boolean
            Dim blnItalic As Boolean
            intColour = MobjDomDoc.queryCommandValue("ForeColor")
            FontDialog1.Color = ColorTranslator.FromHtml(fixColour(intColour))
            intSize = MobjDomDoc.queryCommandValue("FontSize")
            Select Case intSize
                Case 1
                    intSize = 8
                Case 2
                    intSize = 10
                Case 3
                    intSize = 12
                Case 4
                    intSize = 14
                Case 5
                    intSize = 16
                Case 6
                    intSize = 20
                Case Else
                    intSize = 26
            End Select
            strFamily = MobjDomDoc.queryCommandValue("FontName")
            blnBold = MobjDomDoc.queryCommandValue("Bold")
            blnItalic = MobjDomDoc.queryCommandValue("Italic")
            objFontStyle = FontStyle.Regular
            If blnBold Then
                objFontStyle = objFontStyle + FontStyle.Bold
            End If
            If blnItalic Then
                objFontStyle = objFontStyle + FontStyle.Italic
            End If
            objFont = New Drawing.Font(strFamily, intSize, objFontStyle)
            FontDialog1.Font = objFont
            If FontDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                MobjDomDoc.execCommand("RemoveFormat", False, Nothing)
                strColour = ColorTranslator.ToHtml(FontDialog1.Color)
                MobjDomDoc.execCommand("ForeColor", False, strColour)
                Select Case FontDialog1.Font.Size
                    Case Is < 9
                        strSize = "1"
                    Case Is < 11
                        strSize = "2"
                    Case Is < 13
                        strSize = "3"
                    Case Is < 15
                        strSize = "4"
                    Case Is < 17
                        strSize = "5"
                    Case Is < 21
                        strSize = "6"
                    Case Else
                        strSize = "7"
                End Select
                MobjDomDoc.execCommand("FontSize", False, strSize)
                MobjDomDoc.execCommand("FontName", False, FontDialog1.Font.Name)
                If FontDialog1.Font.Italic = True Then
                    MobjDomDoc.execCommand("Italic", False, Nothing)
                End If
                If FontDialog1.Font.Bold = True Then
                    MobjDomDoc.execCommand("Bold", False, Nothing)
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", tsbtnColour_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        Try
            Dim objDomDoc As IHTMLDocument2
            objDomDoc = WebBrowser1.Document.DomDocument
            'MobjDomDoc = WebBrowser3.Document.DomDocument
            'MobjDomDoc.designMode = "On"
            'objDomDoc.execCommand("SelectAll", False, Nothing)
            'MobjDomDoc.execCommand("SelectAll", False, Nothing)
            'MobjDomDoc.execCommand("Cut", False, Nothing)
            'objDomDoc.execCommand("Copy", False, Nothing)
            'MobjDomDoc.execCommand("Paste", False, Nothing)
            'objDomDoc.execCommand("Unselect", False, Nothing)
            'MobjDomDoc.execCommand("Unselect", False, Nothing)
            MobjDomDoc.bgColor = objDomDoc.bgColor
            MobjDomDoc.fgColor = objDomDoc.fgColor
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", WebBrowser2_DocumentCompleted, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Function fixColour(ByVal intColour As Integer) As Integer
        Try
            'returns BGR not RGB so fix it
            Dim Intermidiate() As Byte = BitConverter.GetBytes(intColour)
            Dim IntermidiateByte As Byte = Intermidiate(0)
            Intermidiate(0) = Intermidiate(2)
            Intermidiate(2) = IntermidiateByte
            fixColour = BitConverter.ToInt32(Intermidiate, 0)
            'End workaround Bug
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", fixColour, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Function

    Private Sub SaveFile()
        Try
            Dim objXMLEl As MSXML2.IXMLDOMElement
            Dim objXMLEl2 As MSXML2.IXMLDOMElement
            If MblnDirty Then
                SavePage(False)
                MblnDirty = False
            End If
            MobjXMLDoc.selectSingleNode("//MediaDirectory").text = tbMediaDirectory.Text
            MobjXMLDoc.selectSingleNode("//Title").text = tbTitle.Text
            MobjXMLDoc.selectSingleNode("//Url").text = tbURL.Text
            MobjXMLDoc.documentElement.setAttribute("id", tbMWTId.Text)
            objXMLEl = MobjXMLDoc.selectSingleNode("//Author")
            objXMLEl.setAttribute("id", tbAuthorId.Text)
            objXMLEl2 = objXMLEl.selectSingleNode("./Name")
            objXMLEl2.text = tbAuthorName.Text
            objXMLEl2 = objXMLEl.selectSingleNode("./Url")
            objXMLEl2.text = tbAuthorURL.Text
            objXMLEl = MobjXMLDoc.selectSingleNode("//Settings")
            If cbAutoSetPageWhenSeen.Checked Then
                objXMLEl.selectSingleNode("AutoSetPageWhenSeen").text = "true"
            Else
                objXMLEl.selectSingleNode("AutoSetPageWhenSeen").text = "false"
            End If

            saveXml(MobjXMLDoc, TextBox1.Text)
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnSaveFile_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Function getAttribute(ByVal objXMLEL As MSXML2.IXMLDOMElement, ByVal strAttName As String) As String
        Dim strAttVal As String = ""
        Try
            Dim objXMLAt As MSXML2.IXMLDOMAttribute
            strAttVal = ""
            objXMLAt = objXMLEL.getAttributeNode(strAttName)
            If Not objXMLAt Is Nothing Then
                strAttVal = objXMLAt.text
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", getAttribute, " & ex.Message & ", " & ex.TargetSite.Name)
        Finally
            getAttribute = strAttVal
        End Try
    End Function

    Private Sub addAttribute(ByVal objXMLElement As MSXML2.IXMLDOMElement, ByVal strAttName As String, ByVal strValue As String)
        Try
            If strValue = "" Then
                objXMLElement.removeAttribute(strAttName)
            Else
                objXMLElement.setAttribute(strAttName, strValue)
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", addAttribute, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnPrevNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevNode.Click
        Try
            If Not TreeViewPages.SelectedNode.PrevNode Is Nothing Then
                TreeViewPages.SelectedNode = TreeViewPages.SelectedNode.PrevNode
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnPrevNode_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnNextNode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNextNode.Click
        Try
            If Not TreeViewPages.SelectedNode.NextNode Is Nothing Then
                TreeViewPages.SelectedNode = TreeViewPages.SelectedNode.NextNode
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnNextNode_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub saveXml(ByRef objXMLDom As MSXML2.IXMLDOMDocument2, ByVal strFileName As String)
        Try
            Dim objDoc As New Xml.XmlDocument
            Dim writer As System.Xml.XmlTextWriter = New System.Xml.XmlTextWriter(strFileName, Nothing)
            writer.Formatting = Xml.Formatting.Indented
            objDoc.LoadXml(objXMLDom.xml)
            objDoc.Save(writer)
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", saveXml, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MoveRow(ByVal i As Integer)
        Try
            If (DataGridView1.SelectedCells.Count > 0) Then
                Dim curr_index As Integer = DataGridView1.CurrentCell.RowIndex
                Dim curr_col_index As Integer = DataGridView1.CurrentCell.ColumnIndex
                Dim curr_row As DataGridViewRow = DataGridView1.CurrentRow
                DataGridView1.Rows.Remove(curr_row)
                DataGridView1.Rows.Insert(curr_index + i, curr_row)
                DataGridView1.CurrentCell = DataGridView1(curr_col_index, curr_index + i)
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", MoveRow, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnRowUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRowUp.Click
        MoveRow(-1)
    End Sub

    Private Sub btnMoveDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMoveDown.Click
        MoveRow(1)
    End Sub

    Private Function GrammarCheck(ByVal TextToCheck As String) As String

        Dim strReturn As String = TextToCheck
        Try
            ' Create Word and temporary document objects.
            Dim objWord As Microsoft.Office.Interop.Word.Application
            Dim objTempDoc As Microsoft.Office.Interop.Word.Document
            ' If there is no data to spell check, then exit sub here.
            If TextToCheck = "" Then
                Return ""
            End If

            objWord = New Microsoft.Office.Interop.Word.Application
            objTempDoc = objWord.Documents.Add
            objWord.Visible = True

            Dim para As Paragraph = objTempDoc.Paragraphs.Add()
            para.Range.Text = TextToCheck

            objTempDoc.CheckGrammar()
            strReturn = objTempDoc.Range.Text
            objTempDoc.Saved = True
            objTempDoc.Close()

            objWord.Quit()

            ' Microsoft Word must be installed. 
        Catch COMExcep As COMException
            MessageBox.Show( _
                "Microsoft Word must be installed for Spell/Grammar Check " _
                & "to run.", "Spell Checker")

        Catch Excep As Exception
            MessageBox.Show("An error has occured.", "Spell Checker")

        End Try
        Return strReturn
    End Function

    Private Sub PageExists(ByVal objXMLNodes As MSXML2.IXMLDOMNodeList, ByVal strPages(,) As String, ByVal intPages As Integer, ByRef strOutput As String)
        Dim intloop As Integer
        Dim objXMLElement As MSXML2.IXMLDOMElement
        Dim strTarget As String
        Dim strPageName As String
        Dim blnFound As Boolean
        Dim strNode As String
        Dim strPage As String
        Dim strButton As String
        Dim strTemp() As String

        Try
            'loop through the nodes passed and check the target page exists
            For intloop = objXMLNodes.length - 1 To 0 Step -1
                objXMLElement = objXMLNodes.item(intloop)
                strNode = objXMLElement.nodeName
                'if it is a button we want to display the button text in the error
                If strNode = "Button" Then
                    strButton = objXMLElement.text
                Else
                    strButton = ""
                End If
                'get the name of the page the node is in
                strPage = getAttribute(objXMLElement.parentNode, "id")
                'get the target page
                strTarget = getAttribute(objXMLElement, "target")
                If strTarget <> "" Then
                    'If target and page name are the same warn it self refernces (this may be intentional when it is a delay)
                    If strTarget = strPage Then
                        strOutput = strOutput & "<tr><td>" & strPage & "</td><td>" & strNode & " " & strButton & "</td><td><font color=""olive"">Targets its own page</font></td></tr>"
                    Else
                        'get random pages will return an array of page names 
                        'if strTarget does not contain random pages it returns a 1 element array containing strTarget
                        strTemp = GetRandomPages(strTarget)
                        For i = 0 To strTemp.Length - 1
                            strPageName = strTemp(i)
                            blnFound = False
                            'loop through the page array and set the found flag to Y if found
                            For intLoop2 = 0 To intPages
                                If strPages(0, intLoop2) = strPageName Then
                                    strPages(1, intLoop2) = "Y"
                                    blnFound = True
                                    Exit For
                                End If
                            Next
                            'if not found generate an error
                            If Not blnFound Then
                                strOutput = strOutput & "<tr><td>" & strPage & "</td><td>" & strNode & " " & strButton & "</td><td><font color=""red"">Page not found " & strPageName & "</font></td></tr>"
                            End If
                        Next
                    End If
                Else
                    'if no target specified and it is a button or a delay, generate an error
                    Select Case strNode
                        Case "Button", "Delay"
                            strOutput = strOutput & "<tr><td>" & strPage & "</td><td>" & strNode & " " & strButton & "</td><td><font color=""red"">No target page specified</font></td></tr>"
                    End Select
                End If
            Next
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", PageExists, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Function GetRandomPages(ByVal strTarget As String) As String()
        Dim strPre As String
        Dim strPost As String
        Dim strMin As String
        Dim strMax As String
        Dim intPos1 As Integer
        Dim intPos2 As Integer
        Dim intPos3 As Integer
        Dim intMin As Integer
        Dim intMax As Integer
        Dim strPages(0) As String

        strPages(0) = ""
        Try
            strPre = ""
            strPost = ""
            intPos1 = strTarget.IndexOf("(")
            'random pages so split out the static bits either side and the random numbers
            If intPos1 > -1 Then
                intPos2 = strTarget.IndexOf("..", intPos1)
                If (intPos2 > -1) Then
                    intPos3 = strTarget.IndexOf(")", intPos2)
                    If (intPos3 > -1) Then
                        intPos1 = intPos1 + 1
                        intPos2 = intPos2 + 2
                        intPos3 = intPos3 + 1
                        strMin = strTarget.Substring(intPos1, intPos2 - intPos1 - 2)
                        intMin = Long.Parse(strMin)
                        strMax = strTarget.Substring(intPos2, intPos3 - intPos2 - 1)
                        intMax = Long.Parse(strMax)
                        If (intPos1 > 1) Then
                            strPre = strTarget.Substring(0, intPos1 - 1)
                        Else
                            strPre = ""
                        End If
                        strPost = strTarget.Substring(intPos3)
                        'loop through each random page
                        ReDim strPages(intMax - intMin)
                        For i = intMin To intMax
                            strPages(i - intMin) = strPre & i & strPost
                        Next
                    End If
                End If
            Else
                strPages(0) = strTarget
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", GetRandomPages, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
        Return strPages
    End Function

    Private Sub LoopCheck(ByVal strPage() As String, ByRef strOutput As String)
        Dim objXMLNodes As MSXML2.IXMLDOMNodeList
        Dim objXMLElement As MSXML2.IXMLDOMElement
        Dim objXMLElement2 As MSXML2.IXMLDOMElement
        Dim strTarget As String
        Dim strTemp() As String
        Dim strTemp2() As String
        Dim intSize As Integer
        Dim intLoop As Integer
        Dim intLoop2 As Integer

        Try
            StatusStrip1.Items("ToolStripStatusLabel1").Text = strPage(0)
            System.Windows.Forms.Application.DoEvents()
            If strPage.Length < MintLoopCheckDepth Then
                If strPage.Length > 1 And strPage(strPage.Length - 1) = strPage(0) Then
                    strOutput = strOutput & "<tr><td>" & String.Join(", ", strPage) & "</td></tr>"
                Else
                    objXMLElement = MobjXMLPages.selectSingleNode("./Page[@id=""" & strPage(strPage.Length - 1) & """]")
                    'if page is missing ignore as we test for missing pages earlier
                    If Not objXMLElement Is Nothing Then
                        objXMLNodes = objXMLElement.selectNodes(".//Button")
                        For intLoop = objXMLNodes.length - 1 To 0 Step -1
                            objXMLElement2 = objXMLNodes.item(intLoop)
                            strTarget = getAttribute(objXMLElement2, "target")
                            If strTarget <> "" Then
                                strTemp2 = GetRandomPages(strTarget)
                                For intLoop2 = 0 To strTemp2.Length - 1
                                    strTemp = strPage
                                    intSize = strPage.Length
                                    ReDim Preserve strTemp(intSize)
                                    strTemp(intSize) = strTemp2(intLoop2)
                                    LoopCheck(strTemp, strOutput)
                                Next
                            End If
                        Next
                        objXMLNodes = objXMLElement.selectNodes(".//Delay")
                        For intLoop = objXMLNodes.length - 1 To 0 Step -1
                            objXMLElement2 = objXMLNodes.item(intLoop)
                            strTarget = getAttribute(objXMLElement2, "target")
                            If strTarget <> "" Then
                                strTemp2 = GetRandomPages(strTarget)
                                For intLoop2 = 0 To strTemp2.Length - 1
                                    strTemp = strPage
                                    intSize = strPage.Length
                                    ReDim Preserve strTemp(intSize)
                                    strTemp(intSize) = strTemp2(intLoop2)
                                    LoopCheck(strTemp, strOutput)
                                Next
                            End If
                        Next
                        objXMLNodes = objXMLElement.selectNodes(".//Audio")
                        For intLoop = objXMLNodes.length - 1 To 0 Step -1
                            objXMLElement2 = objXMLNodes.item(intLoop)
                            strTarget = getAttribute(objXMLElement2, "target")
                            If strTarget <> "" Then
                                strTemp2 = GetRandomPages(strTarget)
                                For intLoop2 = 0 To strTemp2.Length - 1
                                    strTemp = strPage
                                    intSize = strPage.Length
                                    ReDim Preserve strTemp(intSize)
                                    strTemp(intSize) = strTemp2(intLoop2)
                                    LoopCheck(strTemp, strOutput)
                                Next
                            End If
                        Next
                        objXMLNodes = objXMLElement.selectNodes(".//Video")
                        For intLoop = objXMLNodes.length - 1 To 0 Step -1
                            objXMLElement2 = objXMLNodes.item(intLoop)
                            strTarget = getAttribute(objXMLElement2, "target")
                            If strTarget <> "" Then
                                strTemp2 = GetRandomPages(strTarget)
                                For intLoop2 = 0 To strTemp2.Length - 1
                                    strTemp = strPage
                                    intSize = strPage.Length
                                    ReDim Preserve strTemp(intSize)
                                    strTemp(intSize) = strTemp2(intLoop2)
                                    LoopCheck(strTemp, strOutput)
                                Next
                            End If
                        Next
                    End If
                End If
            End If
            StatusStrip1.Items("ToolStripStatusLabel1").Text = ""
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", LoopCheck, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MediaExists(ByVal objXMLNodes As MSXML2.IXMLDOMNodeList, ByRef strOutput As String)
        Dim intloop As Integer
        Dim objXMLElement As MSXML2.IXMLDOMElement
        Dim strMedia As String
        Dim strNode As String
        Dim strPage As String

        Try
            'loop through the nodes passed and check the media exists
            For intloop = objXMLNodes.length - 1 To 0 Step -1
                objXMLElement = objXMLNodes.item(intloop)
                strNode = objXMLElement.nodeName
                'get the name of the page the node is in
                strPage = getAttribute(objXMLElement.parentNode, "id")
                Select Case strNode
                    Case "Image", "Audio", "Video"
                        strMedia = getAttribute(objXMLElement, "id")
                        If strMedia.IndexOf("*") < 0 Then
                            If System.IO.File.Exists(OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text & "\" & strMedia) Then
                            Else
                                strOutput = strOutput & "<tr><td>" & strPage & "</td><td><font color=""red"">" & strNode & " " & strMedia & "</font></td></tr>"
                            End If
                        End If
                End Select
            Next
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", MediaExists, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuFileSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuFileSave.Click
        SaveFile()
    End Sub

    Private Sub MenuFileLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuFileLoad.Click
        Try
            Dim strNewFile As String
            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                'create backup
                If System.IO.File.Exists(OpenFileDialog1.FileName) = True And blnBackup Then
                    strNewFile = OpenFileDialog1.FileName & "." & Now().ToString("yyyyMMddhhmmss") & ".bkp"
                    System.IO.File.Copy(OpenFileDialog1.FileName, strNewFile, True)
                End If
                loadFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", BtnFile_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuPageSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageSave.Click
        SavePage()
    End Sub

    Private Sub MenuFileNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuFileNew.Click
        Try
            Dim DialogBox As New TextDialog()
            Dim strName As String
            Dim objXMLRoot As MSXML2.IXMLDOMElement
            Dim objXMLEl As MSXML2.IXMLDOMElement
            Dim objXMLEl2 As MSXML2.IXMLDOMElement
            Dim objXMLEl3 As MSXML2.IXMLDOMElement
            DialogBox.Text = "File Name"
            If DialogBox.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                strName = DialogBox.TextBox1.Text
                If strName.IndexOf(".xml") = -1 Then
                    strName = strName & ".xml"
                End If
                TextBox1.Text = txtDirectory & "\" & OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & strName
                objXMLRoot = MobjXMLDoc.createElement("Tease")
                objXMLRoot.setAttribute("scriptVersion", "v0.1")
                objXMLRoot.setAttribute("id", "")

                objXMLEl = MobjXMLDoc.createElement("Title")
                objXMLRoot.appendChild(objXMLEl)

                objXMLEl = MobjXMLDoc.createElement("Url")
                objXMLRoot.appendChild(objXMLEl)

                objXMLEl = MobjXMLDoc.createElement("Author")
                objXMLEl.setAttribute("id", "")
                objXMLEl2 = MobjXMLDoc.createElement("Name")
                objXMLEl.appendChild(objXMLEl2)
                objXMLEl2 = MobjXMLDoc.createElement("Url")
                objXMLEl.appendChild(objXMLEl2)
                objXMLRoot.appendChild(objXMLEl)

                objXMLEl = MobjXMLDoc.createElement("MediaDirectory")
                objXMLRoot.appendChild(objXMLEl)

                objXMLEl = MobjXMLDoc.createElement("Settings")
                objXMLRoot.appendChild(objXMLEl)
                objXMLEl2 = MobjXMLDoc.createElement("AutoSetPageWhenSeen")
                objXMLEl2.text = "false"
                objXMLEl.appendChild(objXMLEl2)

                objXMLEl = MobjXMLDoc.createElement("Pages")
                objXMLRoot.appendChild(objXMLEl)
                objXMLEl2 = MobjXMLDoc.createElement("Page")
                objXMLEl2.setAttribute("id", "start")
                objXMLEl3 = MobjXMLDoc.createElement("Text")
                objXMLEl3.text = "Start Page"
                objXMLEl2.appendChild(objXMLEl3)
                objXMLEl.appendChild(objXMLEl2)
                objXMLRoot.appendChild(objXMLEl)

                MobjXMLDoc.documentElement = objXMLRoot
                saveXml(MobjXMLDoc, TextBox1.Text)
                OpenFileDialog1.FileName = TextBox1.Text
                loadFile(TextBox1.Text)
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnNewFile_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuFileExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuFileExit.Click
        Me.Close()
    End Sub

    Private Sub MenuPageNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageNew.Click
        Try
            Dim objXMLText As MSXML2.IXMLDOMElement
            Dim strText As String
            Dim strHtml As String
            Dim DialogBox As New TextDialog()
            Dim blnDoit As Boolean

            blnDoit = True
            If MblnDirty Then
                Select Case MsgBox("Do you want to saved changes to the current page?" & vbCrLf & "Select Yes to save and create a new page, " & vbCrLf & "No lose changes or " & vbCrLf & "Cancel to to stay on this page", MsgBoxStyle.YesNoCancel, "Unsaved Changes")
                    Case MsgBoxResult.Yes
                        SavePage(False)
                        MblnDirty = False
                    Case MsgBoxResult.Cancel
                        blnDoit = False
                End Select
            End If

            If blnDoit Then
                If DialogBox.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    MstrPage = DialogBox.TextBox1.Text
                    lblPage.Text = MstrPage
                    MobjXMLPage = MobjXMLDoc.createElement("Page")
                    MobjXMLPage.setAttribute("id", MstrPage)
                    MobjXMLPages.appendChild(MobjXMLPage)
                    tbImage.Text = ""
                    If Not PictureBox1.Image Is Nothing Then
                        PictureBox1.Image.Dispose()
                        PictureBox1.Image = Nothing
                    End If
                    If Not PictureBox1.Image Is Nothing Then
                        PictureBox2.Image.Dispose()
                        PictureBox2.Image = Nothing
                    End If
                    objXMLText = MobjXMLDoc.createElement("Text")
                    MobjXMLPage.appendChild(objXMLText)
                    strText = ""
                    strHtml = MstrHtmlTemplate.Replace("[TEXT]", strText)
                    WebBrowser1.DocumentText = strHtml
                    Do While Not WebBrowser1.ReadyState = WebBrowserReadyState.Complete
                        System.Windows.Forms.Application.DoEvents()
                    Loop
                    tbPageText.Text = WebBrowser1.Document.Body.InnerHtml
                    'txtRawText.Text = WebBrowser1.Document.Body.InnerHtml
                    cbDelay.Checked = False
                    tbDelaySeconds.Text = ""
                    tbDelayTarget.Text = ""
                    tbDelayStartWith.Text = ""
                    rbHidden.Checked = False
                    rbNormal.Checked = False
                    rbSecret.Checked = False
                    btnDelay.Tag = ""
                    btnDelay.Enabled = False
                    lblTimer.Text = ""
                    cbMetronome.Checked = False
                    tbMetronome.Text = ""
                    tbAudio.Text = ""
                    tbAudioTarget.Text = ""
                    tbAudioStartAt.Text = ""
                    tbAudioStopAt.Text = ""
                    btnPlayAudio.Enabled = False
                    tbVideo.Text = ""
                    tbVideoTarget.Text = ""
                    tbVideoStartAt.Text = ""
                    tbVideoStopAt.Text = ""
                    DataGridView1.Rows.Clear()
                    For intloop = MobjButtons.GetUpperBound(0) To 1 Step -1
                        FlowLayoutPanel1.Controls.Remove(MobjButtons(intloop))
                        MobjButtons(intloop).Dispose()
                    Next
                    MblnDirty = False
                    PopPageTree()
                    TreeViewPages.SelectedNode = TreeViewPages.Nodes.Item(MstrPage)
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnNewPage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuPageDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageDelete.Click
        Try
            If MsgBox("Are you sure you want to delete " & MstrPage & "?", MsgBoxStyle.YesNo, "Delete Page") = MsgBoxResult.Yes Then
                MobjXMLPage = MobjXMLPages.selectSingleNode("./Page[@id=""" & MstrPage & """]")
                MobjXMLPages.removeChild(MobjXMLPage)
                saveXml(MobjXMLDoc, TextBox1.Text)
                PopPageTree()
                TreeViewPages.SelectedNode = TreeViewPages.Nodes(0)
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnDeletePage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuPageCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageCopy.Click
        Try
            Dim DialogBox As New TextDialog()
            Dim blnDoit As Boolean
            Dim objXMLPage1 As MSXML2.IXMLDOMElement
            objXMLPage1 = MobjXMLPage

            blnDoit = True
            If MblnDirty Then
                Select Case MsgBox("Do you want to saved changes to the current page?" & vbCrLf & "Select Yes to save and create a new page, " & vbCrLf & "No lose changes or " & vbCrLf & "Cancel to to stay on this page", MsgBoxStyle.YesNoCancel, "Unsaved Changes")
                    Case MsgBoxResult.Yes
                        SavePage(False)
                        MblnDirty = False
                    Case MsgBoxResult.Cancel
                        blnDoit = False
                End Select
            End If

            If blnDoit Then
                If DialogBox.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    MstrPage = DialogBox.TextBox1.Text
                    MobjXMLPage = MobjXMLDoc.createElement("Page")
                    MobjXMLPage.setAttribute("id", MstrPage)
                    MobjXMLPages.insertBefore(MobjXMLPage, objXMLPage1.nextSibling)
                    lblPage.Text = MstrPage
                    PopPageTree()
                    SavePage(False)
                    TreeViewPages.SelectedNode = TreeViewPages.Nodes.Item(MstrPage)
                    MblnDirty = False
                    displaypage()
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnCopyPage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuPageSplit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageSplit.Click
        Try
            Dim objXMLPage1 As MSXML2.IXMLDOMElement
            Dim DialogBox As New TextDialog()
            Dim blnDoit As Boolean
            Dim strPage As String

            blnDoit = True
            If MblnDirty Then
                Select Case MsgBox("Do you want to saved changes to the current page?" & vbCrLf & "Select Yes to save and create a new page, " & vbCrLf & "No lose changes or " & vbCrLf & "Cancel to to stay on this page", MsgBoxStyle.YesNoCancel, "Unsaved Changes")
                    Case MsgBoxResult.Yes
                        SavePage(False)
                        MblnDirty = False
                    Case MsgBoxResult.Cancel
                        blnDoit = False
                End Select
            End If

            If blnDoit Then
                If DialogBox.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    strPage = DialogBox.TextBox1.Text
                    Dim objXMLDelay As MSXML2.IXMLDOMElement
                    Dim objXMLButtons As MSXML2.IXMLDOMNodeList
                    Dim objXMLButton As MSXML2.IXMLDOMElement
                    Dim intloop As Integer

                    objXMLPage1 = MobjXMLPage.nextSibling

                    'page set options
                    addAttribute(MobjXMLPage, "set", "")
                    addAttribute(MobjXMLPage, "unset", "")
                    addAttribute(MobjXMLPage, "if-set", "")
                    addAttribute(MobjXMLPage, "if-not-set", "")
                    objXMLDelay = MobjXMLPage.selectSingleNode("./Delay")
                    If Not objXMLDelay Is Nothing Then
                        MobjXMLPage.removeChild(objXMLDelay)
                    End If
                    objXMLButtons = MobjXMLPage.selectNodes("./Button")
                    For intloop = objXMLButtons.length - 1 To 0 Step -1
                        objXMLButton = objXMLButtons.item(intloop)
                        MobjXMLPage.removeChild(objXMLButton)
                    Next
                    objXMLButton = MobjXMLDoc.createElement("Button")
                    objXMLButton.setAttribute("target", strPage)
                    objXMLButton.text = "Continue"
                    MobjXMLPage.appendChild(objXMLButton)

                    MstrPage = DialogBox.TextBox1.Text
                    MobjXMLPage = MobjXMLDoc.createElement("Page")
                    MobjXMLPage.setAttribute("id", MstrPage)
                    MobjXMLPages.insertBefore(MobjXMLPage, objXMLPage1.nextSibling)
                    lblPage.Text = MstrPage
                    PopPageTree()
                    SavePage(False)
                    TreeViewPages.SelectedNode = TreeViewPages.Nodes.Item(MstrPage)
                    MblnDirty = False
                    displaypage()
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnCopyPage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuPageDownLoadError_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageDownLoadError.Click
        Try
            Dim frmDialogue As New ErrorPopUp
            frmDialogue.txtComment.Text = MstrComment
            frmDialogue.txtError.Text = MstrError
            frmDialogue.ShowDialog()
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnErrors_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuPageDeleteError_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageDeleteError.Click
        Try
            Dim objXMLError As MSXML2.IXMLDOMElement
            Dim objXMLComment As MSXML2.IXMLDOMNode
            objXMLError = MobjXMLPage.selectSingleNode("./Errors")
            If Not objXMLError Is Nothing Then
                MobjXMLPage.removeChild(objXMLError)
            End If
            For Each objXMLComment In MobjXMLPage.childNodes
                MobjXMLPage.removeChild(objXMLComment)
            Next
            SavePage(True)
            PopPageTree()
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnDelError_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuToolsUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsUpload.Click
        Try
            Dim objXMLMediaFiles As MSXML2.IXMLDOMNodeList
            Dim objXMLMediaFile As MSXML2.IXMLDOMElement
            Dim intLoop As Integer
            Dim intLoop2 As Integer
            Dim strMediaFile As String
            Dim strUpload As String
            Dim strMediaFileDir As String
            Dim strUploadDir As String
            Dim strFiles() As String
            strMediaFileDir = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text
            strUploadDir = strMediaFileDir & "\Upload"
            If Not System.IO.Directory.Exists(strUploadDir) Then
                System.IO.Directory.CreateDirectory(strUploadDir)
            End If
            strFiles = System.IO.Directory.GetFiles(strUploadDir)
            If Not strFiles Is Nothing Then
                For intLoop = 0 To strFiles.Length - 1
                    System.IO.File.Delete(strFiles(intLoop))
                Next
            End If

            'Images
            objXMLMediaFiles = MobjXMLDoc.selectNodes("//Image")
            For intLoop = 0 To objXMLMediaFiles.length - 1
                objXMLMediaFile = objXMLMediaFiles.item(intLoop)
                strMediaFile = strMediaFileDir & "\" & getAttribute(objXMLMediaFile, "id")
                If strMediaFile.IndexOf("*") > -1 Then
                    strFiles = System.IO.Directory.GetFiles(strMediaFileDir, getAttribute(objXMLMediaFile, "id"))
                    If Not strFiles Is Nothing Then
                        For intLoop2 = 0 To strFiles.Length - 1
                            strUpload = strUploadDir & strFiles(intLoop2).Substring(strFiles(intLoop2).LastIndexOf("\"))
                            If System.IO.File.Exists(strFiles(intLoop2)) And Not System.IO.File.Exists(strUpload) Then
                                System.IO.File.Copy(strFiles(intLoop2), strUpload)
                            End If
                        Next
                    End If
                Else
                    strUpload = strUploadDir & "\" & getAttribute(objXMLMediaFile, "id")
                    If System.IO.File.Exists(strMediaFile) And Not System.IO.File.Exists(strUpload) Then
                        System.IO.File.Copy(strMediaFile, strUpload)
                    End If
                End If
            Next

            'Audio
            objXMLMediaFiles = MobjXMLDoc.selectNodes("//Audio")
            For intLoop = 0 To objXMLMediaFiles.length - 1
                objXMLMediaFile = objXMLMediaFiles.item(intLoop)
                strMediaFile = strMediaFileDir & "\" & getAttribute(objXMLMediaFile, "id")
                If strMediaFile.IndexOf("*") > -1 Then
                    strFiles = System.IO.Directory.GetFiles(strMediaFileDir, getAttribute(objXMLMediaFile, "id"))
                    If Not strFiles Is Nothing Then
                        For intLoop2 = 0 To strFiles.Length - 1
                            strUpload = strUploadDir & strFiles(intLoop2).Substring(strFiles(intLoop2).LastIndexOf("\"))
                            If System.IO.File.Exists(strFiles(intLoop2)) And Not System.IO.File.Exists(strUpload) Then
                                System.IO.File.Copy(strFiles(intLoop2), strUpload)
                            End If
                        Next
                    End If
                Else
                    strUpload = strUploadDir & "\" & getAttribute(objXMLMediaFile, "id")
                    If System.IO.File.Exists(strMediaFile) And Not System.IO.File.Exists(strUpload) Then
                        System.IO.File.Copy(strMediaFile, strUpload)
                    End If
                End If
            Next

            'Video
            objXMLMediaFiles = MobjXMLDoc.selectNodes("//Video")
            For intLoop = 0 To objXMLMediaFiles.length - 1
                objXMLMediaFile = objXMLMediaFiles.item(intLoop)
                strMediaFile = strMediaFileDir & "\" & getAttribute(objXMLMediaFile, "id")
                If strMediaFile.IndexOf("*") > -1 Then
                    strFiles = System.IO.Directory.GetFiles(strMediaFileDir, getAttribute(objXMLMediaFile, "id"))
                    If Not strFiles Is Nothing Then
                        For intLoop2 = 0 To strFiles.Length - 1
                            strUpload = strUploadDir & strFiles(intLoop2).Substring(strFiles(intLoop2).LastIndexOf("\"))
                            If System.IO.File.Exists(strFiles(intLoop2)) And Not System.IO.File.Exists(strUpload) Then
                                System.IO.File.Copy(strFiles(intLoop2), strUpload)
                            End If
                        Next
                    End If
                Else
                    strUpload = strUploadDir & "\" & getAttribute(objXMLMediaFile, "id")
                    If System.IO.File.Exists(strMediaFile) And Not System.IO.File.Exists(strUpload) Then
                        System.IO.File.Copy(strMediaFile, strUpload)
                    End If
                End If
            Next
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnUpLoad_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuToolsSort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsSort.Click
        Dim blnFinished As Boolean
        Dim intLoop As Integer
        Dim objXMLPage1 As MSXML2.IXMLDOMElement
        Dim objXMLPage2 As MSXML2.IXMLDOMElement
        Dim strId1 As String
        Dim strId2 As String
        Me.Cursor = Cursors.WaitCursor
        Try
            Do
                blnFinished = True
                For intLoop = 0 To MobjXMLPages.childNodes.length - 2
                    If MobjXMLPages.childNodes(intLoop).nodeType = MSXML2.DOMNodeType.NODE_ELEMENT Then
                        objXMLPage1 = MobjXMLPages.childNodes(intLoop)
                        strId1 = getAttribute(objXMLPage1, "id")
                        If objXMLPage1.nextSibling.nodeType = MSXML2.DOMNodeType.NODE_ELEMENT Then
                            objXMLPage2 = objXMLPage1.nextSibling
                            strId2 = getAttribute(objXMLPage2, "id")
                            If Not strId1.ToLower = "start" Then
                                If strId2.ToLower = "start" Then
                                    blnFinished = False
                                    MobjXMLPages.insertBefore(objXMLPage2, objXMLPage1)
                                Else
                                    If strId2.ToLower < strId1.ToLower Then
                                        blnFinished = False
                                        MobjXMLPages.insertBefore(objXMLPage2, objXMLPage1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                If blnFinished Then
                    Exit Do
                End If
            Loop
            saveXml(MobjXMLDoc, TextBox1.Text)
            PopPageTree()
            TreeViewPages.SelectedNode = TreeViewPages.Nodes(0)
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", MenuToolsSort_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub MenuToolsCheck_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsCheck.Click
        Dim objXMLNodes As MSXML2.IXMLDOMNodeList
        Dim objXMLElement As MSXML2.IXMLDOMElement
        Dim objXMLElement2 As MSXML2.IXMLDOMElement
        Dim strTarget As String
        Dim strPage As String
        Dim strDelay As String
        Dim intDelay As Integer
        Dim strOutput As String
        Dim intLoop As Integer
        Dim intLoop2 As Integer
        Dim intPages As Integer
        Dim blnFound As Boolean
        Dim strPages(1, 0) As String
        Dim strTemp(0) As String
        Dim blnTargetNotFound As Boolean
        Dim intPos1 As Integer
        Dim intPos2 As Integer
        Dim intPos3 As Integer
        Dim intMin As Integer
        Dim intMax As Integer
        Dim strMin As String
        Dim strMax As String
        Dim objProgress As System.Windows.Forms.ToolStripProgressBar
        objProgress = StatusStrip1.Items("ToolStripProgressBar1")

        Cursor = Cursors.WaitCursor
        strOutput = ""
        Try
            'Out put is displayed in a webbrowser control so we need to create the html
            strOutput = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN""><html><head><title></title><style type=""text/css"">body { background-color:white; color:#000000; font-family: Tahoma; font-size:12pt; }</style></head><body><table border=""1"">"
            strOutput = strOutput & "<tr><th>Page</th><th>Element</th><th><font color=""red"">Error</font> / <font color=""olive"">Warning</font></th></tr>"
            'create an array of pages so we don't have to search the xml every time
            strPages(0, 0) = ""
            strPages(1, 0) = "N"
            intPages = -1
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page")
            'step through each page
            For intLoop = objXMLNodes.length - 1 To 0 Step -1
                objXMLElement = objXMLNodes.item(intLoop)
                strTarget = getAttribute(objXMLElement, "id")
                blnFound = False
                'loop through the rest of the pages to see if we have a duplicate
                For intLoop2 = 0 To intPages
                    If strPages(0, intLoop2) = strTarget Then
                        blnFound = True
                        strOutput = strOutput & "<tr><td>" & strTarget & "</td><td/><td><font color=""red"">Duplicate Page found</font></td></tr>"
                        Exit For
                    End If
                Next
                'if no duplcate found add it to the page array
                If Not blnFound Then
                    If strPages(0, 0) = "" Then
                        strPages(0, 0) = strTarget
                        intPages = intPages + 1
                    Else
                        intPages = intPages + 1
                        ReDim Preserve strPages(1, intPages)
                        strPages(0, intPages) = strTarget
                        strPages(1, intPages) = "N"
                    End If
                End If
            Next
            'check all the button page targets
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page//Button")
            PageExists(objXMLNodes, strPages, intPages, strOutput)
            'check all the delay page targets
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page//Delay")
            PageExists(objXMLNodes, strPages, intPages, strOutput)
            'check delay is less than the maximum
            For intLoop = objXMLNodes.length - 1 To 0 Step -1
                objXMLElement2 = objXMLNodes.item(intLoop)
                strDelay = getAttribute(objXMLElement2, "seconds")

                intPos1 = strDelay.IndexOf("(")
                'random pages so split out the static bits either side and the random numbers
                If intPos1 > -1 Then
                    intPos2 = strDelay.IndexOf("..", intPos1)
                    If (intPos2 > -1) Then
                        intPos3 = strDelay.IndexOf(")", intPos2)
                        If (intPos3 > -1) Then
                            intPos1 = intPos1 + 1
                            intPos2 = intPos2 + 2
                            intPos3 = intPos3 + 1
                            strMin = strDelay.Substring(intPos1, intPos2 - intPos1 - 2)
                            intMin = Convert.ToInt32(strMin)
                            strMax = strDelay.Substring(intPos2, intPos3 - intPos2 - 1)
                            intMax = Convert.ToInt32(strMax)
                            If intMax > intMin Then
                                intDelay = intMax
                            Else
                                intDelay = intMin
                            End If
                        Else
                            intDelay = 0
                        End If
                    Else
                        intDelay = 0
                    End If
                Else
                    intDelay = Convert.ToInt32(strDelay)
                End If

                If intDelay > MintMaxDelay Then
                    strPage = getAttribute(objXMLElement2.parentNode, "id")
                    strOutput = strOutput & "<tr><td>" & strPage & "</td><td>Delay</td><td><font color=""olive"">Delay " & strDelay & " greater than " & MintMaxDelay.ToString & "</font></td></tr>"
                End If
            Next
            'check all the Audio page targets
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page//Audio")
            PageExists(objXMLNodes, strPages, intPages, strOutput)
            'check all the Video page targets
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page//Video")
            PageExists(objXMLNodes, strPages, intPages, strOutput)
            'Loop through the page array for any pages still set to N
            For intLoop2 = 0 To intPages - 1
                'if it is N this page has not been referenced (don't bother with start as it is the first page)
                If strPages(1, intLoop2) = "N" And strPages(0, intLoop2) <> "start" Then
                    strOutput = strOutput & "<tr><td>" & strPages(0, intLoop2) & "</td><td/><td><font color=""olive"">Page not referenced</font></td></tr>"
                End If
                objXMLElement = MobjXMLPages.selectSingleNode("./Page[@id=""" & strPages(0, intLoop2) & """]")
                blnTargetNotFound = True
                'if page is missing ignore as we test for missing pages earlier
                If Not objXMLElement Is Nothing Then
                    objXMLNodes = objXMLElement.selectNodes(".//Button")
                    For intLoop = objXMLNodes.length - 1 To 0 Step -1
                        objXMLElement2 = objXMLNodes.item(intLoop)
                        strTarget = getAttribute(objXMLElement2, "target")
                        If strTarget <> "" Then
                            blnTargetNotFound = False
                            Exit For
                        End If
                    Next
                    If blnTargetNotFound Then
                        objXMLNodes = objXMLElement.selectNodes(".//Delay")
                        For intLoop = objXMLNodes.length - 1 To 0 Step -1
                            objXMLElement2 = objXMLNodes.item(intLoop)
                            strTarget = getAttribute(objXMLElement2, "target")
                            If strTarget <> "" Then
                                blnTargetNotFound = False
                                Exit For
                            End If
                        Next
                    End If
                    If blnTargetNotFound Then
                        objXMLNodes = objXMLElement.selectNodes(".//Audio")
                        For intLoop = objXMLNodes.length - 1 To 0 Step -1
                            objXMLElement2 = objXMLNodes.item(intLoop)
                            strTarget = getAttribute(objXMLElement2, "target")
                            If strTarget <> "" Then
                                blnTargetNotFound = False
                                Exit For
                            End If
                        Next
                    End If
                    If blnTargetNotFound Then
                        objXMLNodes = objXMLElement.selectNodes(".//Video")
                        For intLoop = objXMLNodes.length - 1 To 0 Step -1
                            objXMLElement2 = objXMLNodes.item(intLoop)
                            strTarget = getAttribute(objXMLElement2, "target")
                            If strTarget <> "" Then
                                blnTargetNotFound = False
                                Exit For
                            End If
                        Next
                    End If
                    If blnTargetNotFound Then
                        strOutput = strOutput & "<tr><td>" & strPages(0, intLoop2) & "</td><td/><td><font color=""olive"">Page has no target</font></td></tr>"
                    End If
                End If
            Next
            strOutput = strOutput & "</table><table border=""1""><tr><th>Page</th><th>Media Missing</th></tr>"
            'check all the images exist
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page//Image")
            MediaExists(objXMLNodes, strOutput)
            'check all the images exist
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page//Audio")
            MediaExists(objXMLNodes, strOutput)
            'check all the images exist
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page//Video")
            MediaExists(objXMLNodes, strOutput)
            strOutput = strOutput & "</table></body></html>"

        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", MenuToolsSort_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
        Dim frmDialogue As New CheckPopUp
        frmDialogue.WBCheck.DocumentText = strOutput
        Do While Not frmDialogue.WBCheck.ReadyState = WebBrowserReadyState.Complete
            System.Windows.Forms.Application.DoEvents()
        Loop
        Cursor = Cursors.Default
        frmDialogue.ShowDialog()
    End Sub


    Private Sub MenuToolsNyx_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsNyx.Click
        Try
            Dim objXMLDomText As New MSXML2.DOMDocument
            Dim objXMLNodes As MSXML2.IXMLDOMNodeList
            Dim objXMLElement As MSXML2.IXMLDOMElement
            Dim objXMLPages As MSXML2.IXMLDOMNode
            Dim objXMLPageNode As MSXML2.IXMLDOMNode
            Dim objXMLPage As MSXML2.IXMLDOMNode
            Dim objXMLImage As MSXML2.IXMLDOMElement
            Dim objXMLText As MSXML2.IXMLDOMElement
            Dim objXMLDelay As MSXML2.IXMLDOMElement
            Dim objXMLMetronome As MSXML2.IXMLDOMElement
            Dim objXMLAudio As MSXML2.IXMLDOMElement
            Dim objXMLVideo As MSXML2.IXMLDOMElement
            Dim objXMLButtons As MSXML2.IXMLDOMNodeList
            Dim objXMLButton As MSXML2.IXMLDOMElement
            Dim strImage As String
            Dim strAudio As String
            Dim strVideo As String
            Dim strText As String
            Dim strSeconds As String
            Dim strTarget As String
            Dim strStyle As String
            Dim strBPM As String
            Dim strButtonTarget As String
            Dim strButtonText As String
            Dim strSet As String
            Dim strUnSet As String
            Dim strIfSet As String
            Dim strIfNotSet As String
            Dim intloop As Integer
            Dim intloop2 As Integer
            Dim strScript As String = ""
            Dim strTxtSplit() As String
            Dim strTemp As String
            Dim intRandomPage As Integer
            Dim strRandomPage As String
            Dim strRandomPages As String
            Dim strButtonSet As String
            Dim strButtonUnSet As String
            Dim strDelaySet As String
            Dim strDelayUnSet As String
            Dim strTmpArr() As String
            Dim strTmpArr2() As String
            Dim strTmpArr3() As String
            Dim strPageName As String
            Dim strPages(1, 0) As String
            Dim intPages As Integer
            Dim blnFound As Boolean
            Dim strNewTarget As String = ""

            objXMLPages = MobjXMLPages.cloneNode(True)
            strPages(0, 0) = ""
            strPages(1, 0) = "N"
            intPages = -1
            'check all the button page targets
            objXMLNodes = objXMLPages.selectNodes("//Page//Button")
            For intloop = objXMLNodes.length - 1 To 0 Step -1
                objXMLElement = objXMLNodes.item(intloop)
                'get the target page
                strTarget = getAttribute(objXMLElement, "target")
                If strTarget <> "" Then
                    'get random pages will return an array of page names 
                    'if strTarget does not contain random pages it returns a 1 element array containing strTarget
                    strTmpArr2 = GetRandomPages(strTarget)
                    strTmpArr3 = EncodeRandomPages(strTarget, strNewTarget)
                    If strTarget <> strNewTarget Then
                        objXMLElement.setAttribute("target", strNewTarget)
                        For i = 0 To strTmpArr2.Length - 1
                            strPageName = strTmpArr2(i)
                            blnFound = False
                            'loop through the rest of the pages to see if we have a duplicate
                            For intloop2 = 0 To intPages
                                If strPages(0, intloop2) = strPageName Then
                                    blnFound = True
                                    Exit For
                                End If
                            Next
                            'if no duplcate found add it to the page array
                            If Not blnFound Then
                                If strPages(0, 0) = "" Then
                                    strPages(0, 0) = strPageName
                                    strPages(1, 0) = strTmpArr3(i)
                                    intPages = intPages + 1
                                Else
                                    intPages = intPages + 1
                                    ReDim Preserve strPages(1, intPages)
                                    strPages(0, intPages) = strPageName
                                    strPages(1, intPages) = strTmpArr3(i)
                                End If
                            End If
                        Next
                    End If
                End If
            Next

            'check all the delay page targets
            objXMLNodes = objXMLPages.selectNodes("//Page//Delay")
            For intloop = objXMLNodes.length - 1 To 0 Step -1
                objXMLElement = objXMLNodes.item(intloop)
                'get the target page
                strTarget = getAttribute(objXMLElement, "target")
                If strTarget <> "" Then
                    'get random pages will return an array of page names 
                    'if strTarget does not contain random pages it returns a 1 element array containing strTarget
                    strTmpArr2 = GetRandomPages(strTarget)
                    strTmpArr3 = EncodeRandomPages(strTarget, strNewTarget)
                    If strTarget <> strNewTarget Then
                        objXMLElement.setAttribute("target", strNewTarget)
                        For i = 0 To strTmpArr2.Length - 1
                            strPageName = strTmpArr2(i)
                            blnFound = False
                            'loop through the rest of the pages to see if we have a duplicate
                            For intloop2 = 0 To intPages
                                If strPages(0, intloop2) = strPageName Then
                                    blnFound = True
                                    Exit For
                                End If
                            Next
                            'if no duplcate found add it to the page array
                            If Not blnFound Then
                                If strPages(0, 0) = "" Then
                                    strPages(0, 0) = strPageName
                                    strPages(1, 0) = strTmpArr3(i)
                                    intPages = intPages + 1
                                Else
                                    intPages = intPages + 1
                                    ReDim Preserve strPages(1, intPages)
                                    strPages(0, intPages) = strPageName
                                    strPages(1, intPages) = strTmpArr3(i)
                                End If
                            End If
                        Next
                    End If
                End If
            Next


            For intloop2 = 0 To intPages
                objXMLElement = objXMLPages.selectSingleNode("//Page[@id='" & strPages(0, intloop2) & "']")
                objXMLElement.setAttribute("id", strPages(1, intloop2))
            Next
            'Still need to change the page names

            intRandomPage = Integer.Parse(txtRandom.Text)
            strRandomPages = ""
            For Each objXMLPageNode In objXMLPages.childNodes
                If objXMLPageNode.nodeType = MSXML2.DOMNodeType.NODE_ELEMENT Then
                    objXMLPage = objXMLPageNode
                    strScript = strScript & getAttribute(objXMLPage, "id") & "#page("
                    'text
                    objXMLText = objXMLPage.selectSingleNode("./Text")
                    strText = objXMLText.xml
                    strText = strText.Replace("<Text>", "")
                    strText = strText.Replace("</Text>", "")
                    strText = strText.Replace("'", "&apos;")
                    strText = strText.Replace("<p>", "")
                    strText = strText.Replace("</p>", "")
                    strText = strText.Replace("<P>", "")
                    strText = strText.Replace("</P>", "")
                    strScript = strScript & "'"
                    strScript = strScript & strText
                    strScript = strScript & "'"
                    'image
                    objXMLImage = objXMLPage.selectSingleNode("Image")
                    If Not objXMLImage Is Nothing Then
                        strImage = getAttribute(objXMLImage, "id")
                        strImage = strImage.Replace("_", "-")
                        strScript = strScript & ",pic(""" & strImage.ToLower & """)"
                    End If

                    'Buttons
                    objXMLButtons = objXMLPage.selectNodes("./Button")
                    'populate with buttons for this page
                    If objXMLButtons.length > 0 Then
                        strScript = strScript & ",vert("
                        strScript = strScript & "buttons("
                    End If
                    For intloop = objXMLButtons.length - 1 To 0 Step -1
                        objXMLButton = objXMLButtons.item(intloop)
                        strButtonSet = getAttribute(objXMLButton, "set")
                        strButtonUnSet = getAttribute(objXMLButton, "unset")
                        strButtonTarget = getAttribute(objXMLButton, "target")
                        If strButtonTarget.IndexOf("(") > -1 Then
                            strButtonTarget = strButtonTarget.Replace("(", "")
                            strButtonTarget = strButtonTarget.Replace(")", "")
                            strTmpArr = strButtonTarget.Split("..")
                            strRandomPage = intRandomPage.ToString
                            intRandomPage = intRandomPage + 1
                            strButtonTarget = strRandomPage & "#"
                            strRandomPages = strRandomPages & strRandomPage & "#page("""",pic(""blank.jpg""),delay(0sec,range(" & strTmpArr(0) & "," & strTmpArr(2) & "),style:secret));" & vbCrLf
                        Else
                            strButtonTarget = strButtonTarget & "#"
                        End If
                        strButtonText = objXMLButton.text
                        If strButtonSet & strButtonUnSet = "" Then
                            strScript = strScript & strButtonTarget & ","
                        Else
                            strRandomPage = intRandomPage.ToString
                            intRandomPage = intRandomPage + 1
                            strScript = strScript & strRandomPage & "#,"
                            strRandomPages = strRandomPages & strRandomPage & "#page("""",pic(""blank.jpg""),delay(0sec," & strButtonTarget & ")"
                            If strButtonSet <> "" Then
                                strTxtSplit = strButtonSet.Split(",")
                                strTemp = ",set("
                                For intloop2 = 0 To strTxtSplit.Length - 1
                                    strTemp = strTemp & strTxtSplit(intloop2) & "#"
                                    If intloop2 > 0 And intloop2 <> strTxtSplit.Length - 1 Then
                                        strTemp = strTemp & ","
                                    End If
                                Next
                                strTemp = strTemp & ")"
                                strRandomPages = strRandomPages & strTemp
                            End If

                            If strButtonUnSet <> "" Then
                                strTxtSplit = strButtonUnSet.Split(",")
                                strTemp = ",unset("
                                For intloop2 = 0 To strTxtSplit.Length - 1
                                    strTemp = strTemp & strTxtSplit(intloop2) & "#"
                                    If intloop2 > 0 And intloop2 <> strTxtSplit.Length - 1 Then
                                        strTemp = strTemp & ","
                                    End If
                                Next
                                strTemp = strTemp & ")"
                                strRandomPages = strRandomPages & strTemp
                            End If
                            strRandomPages = strRandomPages & ");" & vbCrLf
                        End If
                        strScript = strScript & """" & strButtonText & """"
                        If intloop > 0 Then
                            strScript = strScript & ","
                        End If
                        'strButtonIfSet = getAttribute(objXMLButton, "if-set")
                        'strButtonIfNotSet = getAttribute(objXMLButton, "if-not-set")
                    Next
                    If objXMLButtons.length > 0 Then
                        strScript = strScript & "))"
                    End If

                    'delay
                    objXMLDelay = objXMLPage.selectSingleNode("./Delay")
                    'Delay
                    If Not objXMLDelay Is Nothing Then
                        strSeconds = getAttribute(objXMLDelay, "seconds")
                        strTarget = getAttribute(objXMLDelay, "target")
                        If strTarget.IndexOf("(") > -1 Then
                            strTarget = strTarget.Replace("(", "")
                            strTarget = strTarget.Replace(")", "")
                            strTmpArr = strTarget.Split("..")
                            strTarget = "range(" & strTmpArr(0) & "," & strTmpArr(2) & ")"
                        Else
                            strTarget = strTarget & "#"
                        End If
                        strStyle = getAttribute(objXMLDelay, "style")
                        If strScript.Substring(strScript.Length - 1, 1) <> "(" Then
                            strScript = strScript & ","
                        End If
                        strDelaySet = getAttribute(objXMLDelay, "set")
                        strDelayUnSet = getAttribute(objXMLDelay, "unset")
                        If strDelaySet & strDelayUnSet = "" Then
                            strScript = strScript & "delay(" & strSeconds & "sec," & strTarget
                        Else
                            strRandomPage = intRandomPage.ToString
                            intRandomPage = intRandomPage + 1
                            strScript = strScript & "delay(" & strSeconds & "sec," & strRandomPage & "#"
                            strRandomPages = strRandomPages & strRandomPage & "#page("""",pic(""blank.jpg""),delay(0sec," & strTarget & ")"
                            If strDelaySet <> "" Then
                                strTxtSplit = strDelaySet.Split(",")
                                strTemp = ",set("
                                For intloop2 = 0 To strTxtSplit.Length - 1
                                    strTemp = strTemp & strTxtSplit(intloop2) & "#"
                                    If intloop2 > 0 And intloop2 <> strTxtSplit.Length - 1 Then
                                        strTemp = strTemp & ","
                                    End If
                                Next
                                strTemp = strTemp & ")"
                                strRandomPages = strRandomPages & strTemp
                            End If

                            If strDelayUnSet <> "" Then
                                strTxtSplit = strDelayUnSet.Split(",")
                                strTemp = ",unset("
                                For intloop2 = 0 To strTxtSplit.Length - 1
                                    strTemp = strTemp & strTxtSplit(intloop2) & "#"
                                    If intloop2 > 0 And intloop2 <> strTxtSplit.Length - 1 Then
                                        strTemp = strTemp & ","
                                    End If
                                Next
                                strTemp = strTemp & ")"
                                strRandomPages = strRandomPages & strTemp
                            End If
                            strRandomPages = strRandomPages & ");" & vbCrLf
                        End If
                        If strStyle <> "normal" Then
                            strScript = strScript & ",style:" & strStyle
                        End If
                        strScript = strScript & ")"
                        'delay set options
                        'txtDelayIfSet.Text = getAttribute(objXMLDelay, "if-set")
                        'txtDelayIfNotSet.Text = getAttribute(objXMLDelay, "if-not-set")
                    End If

                    'Audio
                    objXMLAudio = objXMLPage.selectSingleNode("./Audio")
                    If Not objXMLAudio Is Nothing Then
                        strAudio = getAttribute(objXMLAudio, "id")
                        If strScript.Substring(strScript.Length - 1, 1) <> "(" Then
                            strScript = strScript & ","
                        End If
                        strScript = strScript & "hidden:sound(id:'" & strAudio.ToLower & "')"
                    End If

                    'strScript = strScript & ")"

                    strSet = getAttribute(objXMLPage, "set")
                    If strSet <> "" Then
                        strTxtSplit = strSet.Split(",")
                        strTemp = ",set("
                        For intloop = 0 To strTxtSplit.Length - 1
                            strTemp = strTemp & strTxtSplit(intloop) & "#"
                            If intloop > 0 And intloop <> strTxtSplit.Length - 1 Then
                                strTemp = strTemp & ","
                            End If
                        Next
                        strTemp = strTemp & ")"
                        strScript = strScript & strTemp
                    End If

                    strUnSet = getAttribute(objXMLPage, "unset")
                    If strUnSet <> "" Then
                        strTxtSplit = strUnSet.Split(",")
                        strTemp = ",unset("
                        For intloop = 0 To strTxtSplit.Length - 1
                            strTemp = strTemp & strTxtSplit(intloop) & "#"
                            If intloop > 0 And intloop <> strTxtSplit.Length - 1 Then
                                strTemp = strTemp & ","
                            End If
                        Next
                        strTemp = strTemp & ")"
                        strScript = strScript & strTemp
                    End If

                    strIfSet = getAttribute(objXMLPage, "if-set")
                    If strIfSet <> "" Then
                        strTxtSplit = strIfSet.Split(",")
                        strTemp = ",must("
                        For intloop = 0 To strTxtSplit.Length - 1
                            strTemp = strTemp & strTxtSplit(intloop) & "#"
                            If intloop > 0 And intloop <> strTxtSplit.Length - 1 Then
                                strTemp = strTemp & ","
                            End If
                        Next
                        strTemp = strTemp & ")"
                        strScript = strScript & strTemp
                    End If

                    strIfNotSet = getAttribute(objXMLPage, "if-not-set")
                    If strIfNotSet <> "" Then
                        strTxtSplit = strIfNotSet.Split(",")
                        strTemp = ",mustnot("
                        For intloop = 0 To strTxtSplit.Length - 1
                            strTemp = strTemp & strTxtSplit(intloop) & "#"
                            If intloop > 0 And intloop <> strTxtSplit.Length - 1 Then
                                strTemp = strTemp & ","
                            End If
                        Next
                        strTemp = strTemp & ")"
                        strScript = strScript & strTemp
                    End If

                    'Metronome
                    objXMLMetronome = objXMLPage.selectSingleNode("./Metronome")
                    If Not objXMLMetronome Is Nothing Then
                        strBPM = getAttribute(objXMLMetronome, "bpm")
                    End If
                    'Video
                    objXMLVideo = objXMLPage.selectSingleNode("./Video")
                    If Not objXMLVideo Is Nothing Then
                        strVideo = getAttribute(objXMLVideo, "id")
                    End If
                    strScript = strScript & ");" & vbCrLf
                End If

            Next
            txtNyxScript.Text = strScript & strRandomPages
            'start#page(
            '        '<TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0">Here is some text</FONT></P></TEXTFORMAT><TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#CC0000" LETTERSPACING="0" KERNING="0">C<FONT COLOR="#CCFF00">o<FONT COLOR="#FFFFFF">l<FONT COLOR="#330066">oured</FONT> text</FONT></FONT></FONT></P></TEXTFORMAT><TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0"><B>Bold</B></FONT></P></TEXTFORMAT><TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0"><U>Underline</U></FONT></P></TEXTFORMAT><TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0"><I>Italic</I></FONT></P></TEXTFORMAT><TEXTFORMAT LEADING="2"><P ALIGN="LEFT"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0">Left</FONT></P></TEXTFORMAT><TEXTFORMAT LEADING="2"><P ALIGN="RIGHT"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0">Right</FONT></P></TEXTFORMAT><TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0">Center</FONT></P></TEXTFORMAT>',
            'pic("fiat-panda-cross-01.jpg"),
            'vert(go(page2#),delay(10sec, page2#))
            ');

            'page2#page(
            '        '<TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0">Random Image</FONT></P></TEXTFORMAT>',
            'pic("*.jpg"),
            'vert(buttons(page2#, "Button1", page4#, "Button2", page5#, "Button3"),delay(1min, page2#,style:secret))
            ');

            'page3#page(
            '        '<TEXTFORMAT LEADING="2"><P ALIGN="CENTER"><FONT FACE="FontSans" SIZE="18" COLOR="#FFFFFF" LETTERSPACING="0" KERNING="0">do you like mud?</FONT></P></TEXTFORMAT>',
            'pic("fiat-panda-cross-widescreen-08.jpg"),
            'vert(yn(page2#,page4#),delay(1hrs, page2#,style:hidden))
            ');
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnNyx_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub tscbTextToVisual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            WebBrowser3.Document.Body.InnerHtml = txtRawText.Text
            Do While Not WebBrowser3.ReadyState = WebBrowserReadyState.Complete
                System.Windows.Forms.Application.DoEvents()
            Loop
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", BtnRawText_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub tscbVisualToText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            txtRawText.Text = WebBrowser3.Document.Body.InnerHtml
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", tscbUpdate_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub tbDelayStartWith_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelayStartWith.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbAudioTarget_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAudioTarget.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbAudioStartAt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAudioStartAt.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbAudioStopAt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAudioStopAt.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbAudio_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAudio.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbVideo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVideo.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbVideoTarget_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVideoTarget.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbVideoStartAt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVideoStartAt.TextChanged
        MblnDirty = True
    End Sub

    Private Sub tbVideoStopAt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVideoStopAt.TextChanged
        MblnDirty = True
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim strChecked As String
        strChecked = GrammarCheck(txtRawText.Text)
        txtRawText.Text = strChecked
    End Sub

    Private Sub ListView1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.MouseEnter

    End Sub

    Private Sub ListView1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.MouseLeave

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        Dim strImage As String
        Try
            If ListView1.SelectedItems.Count > 0 Then
                strImage = ListView1.SelectedItems(0).Text
                strImage = strImage.Substring(txtDirectory.Length + 1)
                strImage = strImage.Replace(tbMediaDirectory.Text, "")
                If strImage.Substring(0, 1) = "\" Then
                    strImage = strImage.Substring(1)
                End If
                tbImage.Text = strImage
                PictureBox1.Load(txtDirectory & "\" & tbMediaDirectory.Text & "\" & strImage)
                PictureBox2.Load(txtDirectory & "\" & tbMediaDirectory.Text & "\" & strImage)
                MblnDirty = True
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", ListView1_SelectedIndexChanged, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Try
            txtRawText.Text = tbPageText.Text
            WebBrowser3.Document.Body.InnerHtml = txtRawText.Text
            Do While Not WebBrowser3.ReadyState = WebBrowserReadyState.Complete
                System.Windows.Forms.Application.DoEvents()
            Loop
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", ToolStripButton2_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        tbPageText.Text = txtRawText.Text
    End Sub

    Private Sub tbPageText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPageText.TextChanged
        MblnDirty = True
        Try
            If Not WebBrowser1.Document.Body Is Nothing Then
                WebBrowser1.Document.Body.InnerHtml = tbPageText.Text
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", tbPageText_TextChanged, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub txtRawText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRawText.TextChanged
        Try
            If MstrEditStatus = "Neither" Then
                MstrEditStatus = "Raw"
                WebBrowser3.Document.Body.InnerHtml = txtRawText.Text
                Do While Not WebBrowser3.ReadyState = WebBrowserReadyState.Complete
                    System.Windows.Forms.Application.DoEvents()
                Loop
                MstrEditStatus = "Neither"
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", BtnRawText_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuToolsOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsOptions.Click
        Dim objDomSettings As New MSXML2.DOMDocument
        Dim objXMLEl As MSXML2.IXMLDOMElement
        Dim frmDialogue As New OptionsPopUp
        Dim intTemp As Integer
        Try
            frmDialogue.txtDefaltDir.Text = txtDirectory
            frmDialogue.cbBackup.Checked = blnBackup
            frmDialogue.tbThumbnailSize.Text = MintThumbnailSize
            frmDialogue.tbLoopCheck.Text = MintLoopCheckDepth
            frmDialogue.tbMaxDelay.Text = MintMaxDelay
            frmDialogue.tbImageFiles.Text = MstrImageFilter
            frmDialogue.tbAudioFiles.Text = MstrAudioFilter
            frmDialogue.tbVideoFiles.Text = MstrVideoFilter
            If frmDialogue.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                txtDirectory = frmDialogue.txtDefaltDir.Text
                blnBackup = frmDialogue.cbBackup.Checked
                intTemp = Convert.ToInt32(frmDialogue.tbThumbnailSize.Text)
                If intTemp > 20 Then
                    MintThumbnailSize = intTemp
                End If
                intTemp = Convert.ToInt32(frmDialogue.tbLoopCheck.Text)
                If intTemp > 1 Then
                    MintLoopCheckDepth = intTemp
                End If
                intTemp = Convert.ToInt32(frmDialogue.tbMaxDelay.Text)
                If intTemp > 1 Then
                    MintMaxDelay = intTemp
                End If
                MstrImageFilter = frmDialogue.tbImageFiles.Text
                MstrAudioFilter = frmDialogue.tbAudioFiles.Text
                MstrVideoFilter = frmDialogue.tbVideoFiles.Text
                objDomSettings.load(My.Application.Info.DirectoryPath & "\Settings.xml")
                objXMLEl = objDomSettings.documentElement
                objXMLEl.setAttribute("directory", txtDirectory)
                objXMLEl.setAttribute("backup", blnBackup)
                objXMLEl.setAttribute("thumbnailsize", MintThumbnailSize)
                objXMLEl.setAttribute("loopcheckdepth", MintLoopCheckDepth)
                objXMLEl.setAttribute("maxdelay", MintMaxDelay)
                objXMLEl.setAttribute("imagefilter", MstrImageFilter)
                objXMLEl.setAttribute("audiofilter", MstrAudioFilter)
                objXMLEl.setAttribute("videofilter", MstrVideoFilter)
                saveXml(objDomSettings, My.Application.Info.DirectoryPath & "\Settings.xml")
                If txtDirectory <> "" Then
                    OpenFileDialog1.InitialDirectory = txtDirectory
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", MenuToolsOptions_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuPageRename_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuPageRename.Click
        Try
            Dim DialogBox As New TextDialog()
            Dim blnDoit As Boolean

            blnDoit = True
            If MblnDirty Then
                Select Case MsgBox("Do you want to saved changes to the current page?" & vbCrLf & "Select Yes to save and create a new page, " & vbCrLf & "No lose changes or " & vbCrLf & "Cancel to to stay on this page", MsgBoxStyle.YesNoCancel, "Unsaved Changes")
                    Case MsgBoxResult.Yes
                        SavePage(False)
                        MblnDirty = False
                    Case MsgBoxResult.Cancel
                        blnDoit = False
                End Select
            End If

            If blnDoit Then
                If DialogBox.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    MstrPage = DialogBox.TextBox1.Text
                    MobjXMLPage.setAttribute("id", MstrPage)
                    lblPage.Text = MstrPage
                    PopPageTree()
                    SavePage(False)
                    TreeViewPages.SelectedNode = TreeViewPages.Nodes.Item(MstrPage)
                    MblnDirty = False
                    displaypage()
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnRenamePage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuToolsRefreshMedia_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsRefreshMedia.Click
        fillListView()
    End Sub

    Private Sub ListView2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView2.SelectedIndexChanged
        Dim strAudio As String
        Try
            If ListView2.SelectedItems.Count > 0 Then
                strAudio = ListView2.SelectedItems(0).Text
                OpenFileDialogImage.FileName = strAudio
                tbAudio.Text = OpenFileDialogImage.SafeFileName
                MblnDirty = True
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", ListView2_SelectedIndexChanged, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub ListView3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView3.SelectedIndexChanged
        Dim strVideo As String
        Try
            If ListView3.SelectedItems.Count > 0 Then
                strVideo = ListView3.SelectedItems(0).Text
                OpenFileDialogImage.FileName = strVideo
                tbVideo.Text = OpenFileDialogImage.SafeFileName
                MblnDirty = True
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", ListView3_SelectedIndexChanged, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub MenuToolsLoops_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsLoops.Click
        Dim objXMLNodes As MSXML2.IXMLDOMNodeList
        Dim objXMLElement As MSXML2.IXMLDOMElement
        Dim strTarget As String
        Dim strOutput As String
        Dim intLoop As Integer
        Dim intLoop2 As Integer
        Dim intPages As Integer
        Dim blnFound As Boolean
        Dim strPages(0) As String
        Dim strTemp(0) As String
        Dim objProgress As System.Windows.Forms.ToolStripProgressBar
        objProgress = StatusStrip1.Items("ToolStripProgressBar1")

        Cursor = Cursors.WaitCursor
        strOutput = ""
        Try
            'Out put is displayed in a webbrowser control so we need to create the html
            strOutput = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN""><html><head><title></title><style type=""text/css"">body { background-color:white; color:#000000; font-family: Tahoma; font-size:12pt; }</style></head><body><table border=""1"">"
            'create an array of pages so we don't have to search the xml every time
            strPages(0) = ""
            intPages = -1
            objXMLNodes = MobjXMLDoc.selectNodes("//Pages//Page")
            'step through each page
            For intLoop = objXMLNodes.length - 1 To 0 Step -1
                objXMLElement = objXMLNodes.item(intLoop)
                strTarget = getAttribute(objXMLElement, "id")
                blnFound = False
                'loop through the rest of the pages to see if we have a duplicate
                For intLoop2 = 0 To intPages
                    If strPages(intLoop2) = strTarget Then
                        blnFound = True
                        Exit For
                    End If
                Next
                'if no duplcate found add it to the page array
                If Not blnFound Then
                    If strPages(0) = "" Then
                        strPages(0) = strTarget
                        intPages = intPages + 1
                    Else
                        intPages = intPages + 1
                        ReDim Preserve strPages(intPages)
                        strPages(intPages) = strTarget
                    End If
                End If
            Next
            strOutput = strOutput & "<table border=""1""><tr><th>Loop Check</th></tr>"
            For intLoop2 = 0 To intPages
                objProgress.Value = intLoop2 / intPages * 100
                strTemp(0) = strPages(intLoop2)
                LoopCheck(strTemp, strOutput)
            Next
            strOutput = strOutput & "</table></body></html>"
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", ListView3_SelectedIndexChanged, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try

        Dim frmDialogue As New CheckPopUp
        frmDialogue.WBCheck.DocumentText = strOutput
        Do While Not frmDialogue.WBCheck.ReadyState = WebBrowserReadyState.Complete
            System.Windows.Forms.Application.DoEvents()
        Loop
        Cursor = Cursors.Default
        frmDialogue.ShowDialog()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedTab.Name = "TabPageView" Then
            TreeViewPages.Focus()
        End If
    End Sub

    Private Function EncodeRandomPages(ByVal strTarget As String, ByRef strEncodedTarget As String) As String()
        Dim strPre As String
        Dim strPost As String
        Dim strMin As String
        Dim strMax As String
        Dim intPos1 As Integer
        Dim intPos2 As Integer
        Dim intPos3 As Integer
        Dim intMin As Integer
        Dim intMax As Integer
        Dim intPrefix As Integer
        Dim strPages(0) As String

        strPages(0) = ""
        Try
            strPre = ""
            strPost = ""
            intPos1 = strTarget.IndexOf("(")
            'random pages so split out the static bits either side and the random numbers
            If intPos1 > -1 Then
                intPos2 = strTarget.IndexOf("..", intPos1)
                If (intPos2 > -1) Then
                    intPos3 = strTarget.IndexOf(")", intPos2)
                    If (intPos3 > -1) Then
                        intPos1 = intPos1 + 1
                        intPos2 = intPos2 + 2
                        intPos3 = intPos3 + 1
                        strMin = strTarget.Substring(intPos1, intPos2 - intPos1 - 2)
                        intMin = Long.Parse(strMin)
                        strMax = strTarget.Substring(intPos2, intPos3 - intPos2 - 1)
                        intMax = Long.Parse(strMax)
                        If (intPos1 > 1) Then
                            strPre = strTarget.Substring(0, intPos1 - 1)
                            intPrefix = Convert.ToInt32(GetEncodedPageName(strPre))
                        Else
                            strPre = ""
                            intPrefix = 0
                        End If
                        strPost = strTarget.Substring(intPos3)
                        If strPost <> "" Then
                            strPre = strPre & strPost
                            intPrefix = Convert.ToInt32(GetEncodedPageName(strPre))
                        End If
                        'loop through each random page
                        ReDim strPages(intMax - intMin)
                        For i = intMin To intMax
                            strPages(i - intMin) = Convert.ToString(intPrefix + i)
                        Next
                        strEncodedTarget = "(" & Convert.ToString(intPrefix + intMin) & ".." & Convert.ToString(intPrefix + intMax) & ")"
                    End If
                End If
            Else
                strPages(0) = strTarget
                strEncodedTarget = strTarget
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", GetRandomPages, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
        Return strPages
    End Function

    Private Function GetEncodedPageName(ByVal strPage As String) As String
        Dim intLoop As Integer
        Dim intLetter As Integer
        Dim intWord As Integer
        Dim strTemp As String
        Dim strNewPage As String = ""
        intWord = 0
        For intLoop = 0 To strPage.Length - 1
            strTemp = strPage.Substring(intLoop, 1)
            intLetter = AscW(strTemp)
            intLetter = intLetter * ((intLoop + 1) * 100)
            intWord = intWord + intLetter
        Next
        intWord = intWord * 100
        strNewPage = intWord
        Return strNewPage
    End Function

    Private Sub MenuToolsNyxText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuToolsNyxText.Click
        Dim intFont As Integer
        Dim intPos As Integer
        Dim intPos2 As Integer
        Dim strComment As String
        Dim strText As String
        Dim strFontSize As String
        Dim strSize As String
        Dim objXMLPage As MSXML2.IXMLDOMNode
        Dim objXMLPageChildNode As MSXML2.IXMLDOMNode
        Dim objXMLPageText As MSXML2.IXMLDOMElement
        Dim objXMLTextChild As MSXML2.IXMLDOMElement
        Dim objXMLComment As MSXML2.IXMLDOMComment
        For Each objXMLPage In MobjXMLPages.childNodes
            If objXMLPage.nodeType <> MSXML2.DOMNodeType.NODE_COMMENT Then
                strComment = ""
                strText = ""
                For Each objXMLPageChildNode In objXMLPage.childNodes
                    Select Case objXMLPageChildNode.nodeType
                        Case MSXML2.DOMNodeType.NODE_COMMENT
                            objXMLComment = objXMLPageChildNode
                            strComment = objXMLComment.text
                        Case MSXML2.DOMNodeType.NODE_ELEMENT
                            If objXMLPageChildNode.nodeName = "Text" Then
                                objXMLPageText = objXMLPageChildNode
                            End If
                    End Select
                Next
                If strComment <> "" Then
                    intPos = strComment.IndexOf("text:'")
                    intPos = intPos + 6
                    intPos2 = strComment.IndexOf("'", intPos)
                    strText = strComment.Substring(intPos, (intPos2 - intPos))
                    If intPos > -1 Then
                        intPos = strText.IndexOf("SIZE=""")
                        If intPos = -1 Then
                            intPos = strText.IndexOf("size=""")
                        End If
                        intPos = intPos + 6
                    End If
                    While intPos <> 5
                        intPos2 = strText.IndexOf("""", intPos)
                        strFontSize = strText.Substring(intPos, intPos2 - intPos)
                        intFont = Convert.ToInt32(strFontSize)
                        strSize = GetFontSize(intFont)
                        strText = strText.Substring(0, intPos) & strSize & strText.Substring(intPos + strFontSize.Length)
                        intPos = strText.IndexOf("SIZE=""", intPos2)
                        If intPos = -1 Then
                            intPos = strText.IndexOf("size=""", intPos2)
                        End If
                        If intPos = -1 Then
                            Exit While
                        End If
                        intPos = intPos + 6
                    End While
                    CleanHtml(strText)
                    For intloop = objXMLPageText.childNodes.length - 1 To 0 Step -1
                        If objXMLPageText.childNodes(intloop).nodeType = MSXML2.DOMNodeType.NODE_TEXT Then
                            objXMLPageText.text = ""
                        Else
                            objXMLTextChild = objXMLPageText.childNodes(intloop)
                            objXMLPageText.removeChild(objXMLTextChild)
                        End If
                    Next
                    objXMLPageText.appendChild(MobjXMLDocFrag.documentElement)
                End If
            End If
        Next
    End Sub

    Private Function GetFontSize(ByVal intFont As Integer) As String
        Dim strSize As String
        Select Case intFont
            Case Is < 11
                strSize = "1"
            Case Is < 13
                strSize = "2"
            Case Is < 17
                strSize = "3"
            Case Is < 21
                strSize = "4"
            Case Else
                strSize = "5"
        End Select
        Return strSize
    End Function

    Private Sub CleanHtml(ByVal strText As String)
        strText = strText.Replace("FACE=""FontSans""", "")
        strText = strText.Replace("COLOR=""#FFFFFF""", "")
        strText = strText.Replace("KERNING=""0""", "")
        strText = strText.Replace("LETTERSPACING=""0""", "")
        If strText.IndexOf("<TEXTFORMAT LEADING=""2"">") > -1 Then
            strText = strText.Replace("<TEXTFORMAT LEADING=""2"">", "")
            strText = strText.Replace("</TEXTFORMAT>", "")
        End If

        MobjXMLDocFrag.loadXML(strText)
        If MobjXMLDocFrag.documentElement Is Nothing Then
            strText = "<DIV>" & strText & "</DIV>"
            MobjXMLDocFrag.loadXML(strText)
            If MobjXMLDocFrag.documentElement Is Nothing Then
                strText = strText.Replace(" ", " ")
            End If
        End If

    End Sub

    Private Sub visToraw()
        Try
            If MstrEditStatus = "Neither" Then
                MstrEditStatus = "Vis"
                txtRawText.Text = WebBrowser3.Document.Body.InnerHtml
                MstrEditStatus = "Neither"
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", tscbUpdate_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub WebBrowser3_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser3.DocumentCompleted
        doc = CType(sender, WebBrowser).Document
    End Sub
    Private Sub Doc_Click(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles doc.Click
        visToraw()
    End Sub

    Private Sub doc_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles doc.MouseMove
        visToraw()
    End Sub
End Class
