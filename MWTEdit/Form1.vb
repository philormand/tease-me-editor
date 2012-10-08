Public Class Form1
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

    Private Sub BtnFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFile.Click
        Try
            Dim strNewFile As String
            If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                'create backup
                If System.IO.File.Exists(OpenFileDialog1.FileName) = True Then
                    strNewFile = OpenFileDialog1.FileName & "." & Now().ToString("yyyyMMddhhmmss") & ".bkp"
                    System.IO.File.Copy(OpenFileDialog1.FileName, strNewFile, True)
                End If
                loadFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", BtnFile_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

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
            cbAutoSetPageWhenSeen.Checked = objXMLEl.selectSingleNode("AutoSetPageWhenSeen").text
            TreeViewPages.SelectedNode = TreeViewPages.Nodes(0)
            Application.DoEvents()
            displaypage()
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", loadFile, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MblnDebug Then
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", Form Closing")
        End If
        MobjLogWriter.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MobjLogWriter = My.Computer.FileSystem.OpenTextFileWriter("ErrorLog.txt", True)
        If MblnDebug Then
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", Form Opening")
        End If
        Try
            OpenFileDialog1.DefaultExt = ".xml"
            OpenFileDialog1.Filter = "XML Files|*.xml"
            Dim strPath As String
            Dim objDomSettings As New MSXML2.DOMDocument
            Dim objXMLEl As MSXML2.IXMLDOMElement
            objDomSettings.load(My.Application.Info.DirectoryPath & "\Settings.xml")
            If objDomSettings.documentElement Is Nothing Then
                objXMLEl = objDomSettings.createElement("Settings")
                objXMLEl.setAttribute("directory", "c:\")
                objDomSettings.documentElement = objXMLEl
                saveXml(objDomSettings, My.Application.Info.DirectoryPath & "\Settings.xml")
            End If
            strPath = getAttribute(objDomSettings.documentElement, "directory")
            txtDefaltDir.Text = strPath
            If strPath <> "" Then
                OpenFileDialog1.InitialDirectory = strPath
            Else
                OpenFileDialog1.InitialDirectory = My.Application.Info.DirectoryPath
            End If
            ReDim MobjButtons(0)
            MblnDirty = False
            AddHandler btnDelay.Click, AddressOf DynamicButtonClick
            WebBrowser3.DocumentText = "<html><head></head><body></body></html>"
            Do While Not WebBrowser3.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
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
            TabPage3.Focus()
            Application.DoEvents()
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
                Application.DoEvents()
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
                    Application.DoEvents()
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
                        savepage(False)
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
            If OpenFileDialogImage.ShowDialog() = Windows.Forms.DialogResult.OK Then
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

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
            If OpenFileDialogImage.ShowDialog() = Windows.Forms.DialogResult.OK Then
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
            If OpenFileDialogImage.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If tbVideo.Text <> OpenFileDialogImage.SafeFileName Then
                    tbVideo.Text = OpenFileDialogImage.SafeFileName
                    MblnDirty = True
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnVideo_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub cbAudio_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAudio.CheckedChanged
        MblnDirty = True
    End Sub

    Private Sub cbVideo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbVideo.CheckedChanged
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

    Private Sub btnSavePage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSavePage.Click
        savepage(True)
        MblnDirty = False
    End Sub

    Private Sub btnNewPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewPage.Click
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
                        savepage(False)
                        MblnDirty = False
                    Case MsgBoxResult.Cancel
                        blnDoit = False
                End Select
            End If

            If blnDoit Then
                If DialogBox.ShowDialog = Windows.Forms.DialogResult.OK Then
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
                        Application.DoEvents()
                    Loop
                    WebBrowser2.DocumentText = strHtml
                    Do While Not WebBrowser2.ReadyState = WebBrowserReadyState.Complete
                        Application.DoEvents()
                    Loop
                    txtRawText.Text = WebBrowser1.Document.Body.InnerHtml
                    cbDelay.Checked = False
                    tbDelaySeconds.Text = ""
                    tbDelayTarget.Text = ""
                    rbHidden.Checked = False
                    rbNormal.Checked = False
                    rbSecret.Checked = False
                    btnDelay.Tag = ""
                    btnDelay.Enabled = False
                    lblTimer.Text = ""
                    cbMetronome.Checked = False
                    tbMetronome.Text = ""
                    cbAudio.Checked = False
                    tbAudio.Text = ""
                    btnPlayAudio.Enabled = False
                    cbVideo.Checked = False
                    tbVideo.Text = ""
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

    Private Sub btnDeletePage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeletePage.Click
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

            objXMLImage = MobjXMLPage.selectSingleNode("Image")
            If Not objXMLImage Is Nothing Then
                objXMLImage.setAttribute("id", tbImage.Text)
            Else
                objXMLImage = MobjXMLDoc.createElement("Image")
                objXMLImage.setAttribute("id", tbImage.Text)
                MobjXMLPage.appendChild(objXMLImage)
            End If
            objXMLText = MobjXMLPage.selectSingleNode("./Text")
            If objXMLText Is Nothing Then
                objXMLText = MobjXMLDoc.createElement("Text")
                MobjXMLPage.appendChild(objXMLText)
            End If
            strStyle = HtmlAsXml(WebBrowser2.Document.Body.InnerHtml)
            strStyle = strStyle.Replace("<BR>", "<BR/>")
            strStyle = strStyle.Replace("&nbsp;", "<BR/>")
            strStyle = strStyle.Trim(" ")
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
            objXMLDelay = MobjXMLPage.selectSingleNode("./Delay")
            If objXMLDelay Is Nothing Then
                If cbDelay.Checked Then
                    objXMLDelay = MobjXMLDoc.createElement("Delay")
                    MobjXMLPage.appendChild(objXMLDelay)
                    objXMLDelay.setAttribute("seconds", tbDelaySeconds.Text)
                    objXMLDelay.setAttribute("target", tbDelayTarget.Text)
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
            objXMLAudio = MobjXMLPage.selectSingleNode("./Audio")
            If objXMLAudio Is Nothing Then
                If cbAudio.Checked Then
                    objXMLAudio = MobjXMLDoc.createElement("Audio")
                    MobjXMLPage.appendChild(objXMLAudio)
                    objXMLAudio.setAttribute("id", tbAudio.Text)
                End If
            Else
                If cbAudio.Checked Then
                    objXMLAudio.setAttribute("id", tbAudio.Text)
                Else
                    MobjXMLPage.removeChild(objXMLAudio)
                End If
            End If
            objXMLVideo = MobjXMLPage.selectSingleNode("./Video")
            If objXMLVideo Is Nothing Then
                If cbVideo.Checked Then
                    objXMLVideo = MobjXMLDoc.createElement("Video")
                    MobjXMLPage.appendChild(objXMLVideo)
                    objXMLVideo.setAttribute("id", tbVideo.Text)
                End If
            Else
                If cbVideo.Checked Then
                    objXMLVideo.setAttribute("id", tbVideo.Text)
                Else
                    MobjXMLPage.removeChild(objXMLVideo)
                End If
            End If
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
            saveXml(MobjXMLDoc, TextBox1.Text)
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
            Dim strAudio As String
            Dim strVideo As String
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

            MstrPage = TreeViewPages.SelectedNode.Text
            lblPage.Text = MstrPage
            lblPageName.Text = MstrPage
            MobjXMLPage = MobjXMLPages.selectSingleNode("./Page[@id=""" & MstrPage & """]")
            'page set options
            txtPageSet.Text = getAttribute(MobjXMLPage, "set")
            txtPageUnSet.Text = getAttribute(MobjXMLPage, "unset")
            txtPageIfSet.Text = getAttribute(MobjXMLPage, "if-set")
            txtPageIfNotSet.Text = getAttribute(MobjXMLPage, "if-not-set")
            txtRawText.Text = ""
            'image
            objXMLImage = MobjXMLPage.selectSingleNode("Image")
            If Not objXMLImage Is Nothing Then
                tbImage.Text = getAttribute(objXMLImage, "id")
                If tbImage.Text = "" Or tbImage.Text.IndexOf("*") > -1 Then
                    If Not PictureBox1.Image Is Nothing Then
                        PictureBox1.Image.Dispose()
                        PictureBox1.Image = Nothing
                    End If
                    If Not PictureBox2.Image Is Nothing Then
                        PictureBox2.Image.Dispose()
                        PictureBox2.Image = Nothing
                    End If
                Else
                    strImage = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text & "\" & getAttribute(objXMLImage, "id")
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
                Application.DoEvents()
            Loop
            WebBrowser2.DocumentText = strHtml
            Do While Not WebBrowser2.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            If Not WebBrowser1.Document.Body Is Nothing Then
                txtRawText.Text = WebBrowser1.Document.Body.InnerHtml
            End If
            'delay
            objXMLDelay = MobjXMLPage.selectSingleNode("./Delay")
            'Delay
            If objXMLDelay Is Nothing Then
                cbDelay.Checked = False
                tbDelaySeconds.Text = ""
                tbDelayTarget.Text = ""
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
                btnDelay.Tag = strTarget
                btnDelay.Enabled = True
                Select Case strStyle
                    Case "normal"
                        rbHidden.Checked = False
                        rbNormal.Checked = True
                        rbSecret.Checked = False
                        If strSeconds.IndexOf("..") > -1 Then
                            lblTimer.Text = strSeconds
                        Else
                            intSeconds = strSeconds
                            intMinutes = intSeconds / 60
                            intSeconds = intSeconds - (intMinutes * 60)
                            lblTimer.Text = Microsoft.VisualBasic.Right("0" & intMinutes, 2) & ":" & Microsoft.VisualBasic.Right("0" & intSeconds, 2)
                        End If
                    Case "hidden"
                        rbHidden.Checked = True
                        rbNormal.Checked = False
                        rbSecret.Checked = False
                        lblTimer.Text = ""
                    Case "secret"
                        rbHidden.Checked = False
                        rbNormal.Checked = False
                        rbSecret.Checked = True
                        lblTimer.Text = "00:00"
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
                cbAudio.Checked = False
                tbAudio.Text = ""
                btnPlayAudio.Enabled = False
            Else
                cbAudio.Checked = True
                strAudio = getAttribute(objXMLAudio, "id")
                tbAudio.Text = strAudio
                btnPlayAudio.Enabled = True
            End If
            'Video
            objXMLVideo = MobjXMLPage.selectSingleNode("./Video")
            If objXMLVideo Is Nothing Then
                cbVideo.Checked = False
                tbVideo.Text = ""
            Else
                cbVideo.Checked = True
                strVideo = getAttribute(objXMLVideo, "id")
                tbVideo.Text = strVideo
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
            objXMLError = MobjXMLPage.selectSingleNode("./Errors")
            If objXMLError Is Nothing Then
                MstrComment = ""
                MstrError = ""
                btnErrors.Enabled = False
                btnDelError.Enabled = False
            Else
                MstrError = objXMLError.xml
                MstrComment = ""
                For Each objXMLComment In MobjXMLPage.childNodes
                    If objXMLComment.nodeType = MSXML2.DOMNodeType.NODE_COMMENT Then
                        MstrComment = MstrComment & objXMLComment.xml
                    End If
                Next
                btnErrors.Enabled = True
                btnDelError.Enabled = True
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
            Dim objFont As Font
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
            objFont = New Font(strFamily, intSize, objFontStyle)
            FontDialog1.Font = objFont
            If FontDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
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

    Private Sub WebBrowser2_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser2.DocumentCompleted
        Try
            Dim objDomDoc As IHTMLDocument2
            objDomDoc = WebBrowser2.Document.DomDocument
            MobjDomDoc = WebBrowser3.Document.DomDocument
            MobjDomDoc.designMode = "On"
            objDomDoc.execCommand("SelectAll", False, Nothing)
            MobjDomDoc.execCommand("SelectAll", False, Nothing)
            MobjDomDoc.execCommand("Cut", False, Nothing)
            objDomDoc.execCommand("Copy", False, Nothing)
            MobjDomDoc.execCommand("Paste", False, Nothing)
            objDomDoc.execCommand("Unselect", False, Nothing)
            MobjDomDoc.execCommand("Unselect", False, Nothing)
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

    Private Sub tscbUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tscbUpdate.Click
        Try
            WebBrowser1.Document.Body.InnerHtml = WebBrowser3.Document.Body.InnerHtml
            Do While Not WebBrowser1.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            WebBrowser2.Document.Body.InnerHtml = WebBrowser3.Document.Body.InnerHtml
            Do While Not WebBrowser2.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            txtRawText.Text = WebBrowser1.Document.Body.InnerHtml
            savepage(True)
            MblnDirty = False
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", tscbUpdate_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnSaveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFile.Click
        Try
            Dim objXMLEl As MSXML2.IXMLDOMElement
            Dim objXMLEl2 As MSXML2.IXMLDOMElement
            If MblnDirty Then
                savepage(False)
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

    Private Sub btnNewFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewFile.Click
        Try
            Dim DialogBox As New TextDialog()
            Dim strName As String
            Dim objXMLRoot As MSXML2.IXMLDOMElement
            Dim objXMLEl As MSXML2.IXMLDOMElement
            Dim objXMLEl2 As MSXML2.IXMLDOMElement
            Dim objXMLEl3 As MSXML2.IXMLDOMElement
            DialogBox.Text = "File Name"
            If DialogBox.ShowDialog = Windows.Forms.DialogResult.OK Then
                strName = DialogBox.TextBox1.Text
                If strName.IndexOf(".xml") = -1 Then
                    strName = strName & ".xml"
                End If
                TextBox1.Text = txtDefaltDir.Text & "\" & OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & strName
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

    Private Sub btnSaveDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveDir.Click
        Try
            Dim strPath As String
            Dim objDomSettings As New MSXML2.DOMDocument
            Dim objXMLEl As MSXML2.IXMLDOMElement
            strPath = txtDefaltDir.Text
            objDomSettings.load(My.Application.Info.DirectoryPath & "\Settings.xml")
            objXMLEl = objDomSettings.documentElement
            objXMLEl.setAttribute("directory", strPath)
            saveXml(objDomSettings, My.Application.Info.DirectoryPath & "\Settings.xml")
            If strPath <> "" Then
                OpenFileDialog1.InitialDirectory = strPath
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnSaveDir_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnNyx_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNyx.Click
        Try
            Dim objXMLDomText As New MSXML2.DOMDocument
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
            Dim strHtml As String
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
            Dim intButtons As Integer
            Dim intSeconds As Integer
            Dim intMinutes As Integer
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


            intRandomPage = Integer.Parse(txtRandom.Text)
            strRandomPages = ""
            For Each objXMLPageNode In MobjXMLPages.childNodes
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

    Private Sub btnCopyPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyPage.Click
        Try
            Dim DialogBox As New TextDialog()
            Dim blnDoit As Boolean
            Dim objXMLPage1 As MSXML2.IXMLDOMElement
            objXMLPage1 = MobjXMLPage

            blnDoit = True
            If MblnDirty Then
                Select Case MsgBox("Do you want to saved changes to the current page?" & vbCrLf & "Select Yes to save and create a new page, " & vbCrLf & "No lose changes or " & vbCrLf & "Cancel to to stay on this page", MsgBoxStyle.YesNoCancel, "Unsaved Changes")
                    Case MsgBoxResult.Yes
                        savepage(False)
                        MblnDirty = False
                    Case MsgBoxResult.Cancel
                        blnDoit = False
                End Select
            End If

            If blnDoit Then
                If DialogBox.ShowDialog = Windows.Forms.DialogResult.OK Then
                    MstrPage = DialogBox.TextBox1.Text
                    MobjXMLPage = MobjXMLDoc.createElement("Page")
                    MobjXMLPage.setAttribute("id", MstrPage)
                    MobjXMLPages.insertBefore(MobjXMLPage, objXMLPage1)
                    lblPage.Text = MstrPage
                    PopPageTree()
                    savepage(False)
                    TreeViewPages.SelectedNode = TreeViewPages.Nodes.Item(MstrPage)
                    MblnDirty = False
                    displaypage()
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnCopyPage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnErrors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnErrors.Click
        Try
            Dim frmDialogue As New ErrorPopUp
            frmDialogue.txtComment.Text = MstrComment
            frmDialogue.txtError.Text = MstrError
            frmDialogue.ShowDialog()
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnErrors_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnUpLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpLoad.Click
        Try
            Dim objXMLImages As MSXML2.IXMLDOMNodeList
            Dim objXMLImage As MSXML2.IXMLDOMElement
            Dim intLoop As Integer
            Dim intLoop2 As Integer
            Dim strImage As String
            Dim strUpload As String
            Dim strImageDir As String
            Dim strUploadDir As String
            Dim strFiles() As String
            strImageDir = OpenFileDialog1.FileName.Substring(0, OpenFileDialog1.FileName.LastIndexOf("\") + 1) & tbMediaDirectory.Text
            strUploadDir = strImageDir & "\Upload"
            If Not System.IO.Directory.Exists(strUploadDir) Then
                System.IO.Directory.CreateDirectory(strUploadDir)
            End If
            strFiles = System.IO.Directory.GetFiles(strUploadDir)
            If Not strFiles Is Nothing Then
                For intLoop = 0 To strFiles.Length - 1
                    System.IO.File.Delete(strFiles(intLoop))
                Next
            End If
            objXMLImages = MobjXMLDoc.selectNodes("//Image")
            For intLoop = 0 To objXMLImages.length - 1
                objXMLImage = objXMLImages.item(intLoop)
                tbImage.Text = getAttribute(objXMLImage, "id")
                strImage = strImageDir & "\" & getAttribute(objXMLImage, "id")
                If strImage.IndexOf("*") > -1 Then
                    strFiles = System.IO.Directory.GetFiles(strImageDir, getAttribute(objXMLImage, "id"))
                    If Not strFiles Is Nothing Then
                        For intLoop2 = 0 To strFiles.Length - 1
                            strUpload = strUploadDir & strFiles(intLoop2).Substring(strFiles(intLoop2).LastIndexOf("\"))
                            If System.IO.File.Exists(strFiles(intLoop2)) And Not System.IO.File.Exists(strUpload) Then
                                System.IO.File.Copy(strFiles(intLoop2), strUpload)
                            End If
                        Next
                    End If
                Else
                    strUpload = strUploadDir & "\" & getAttribute(objXMLImage, "id")
                    If System.IO.File.Exists(strImage) And Not System.IO.File.Exists(strUpload) Then
                        System.IO.File.Copy(strImage, strUpload)
                    End If
                End If
            Next
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnUpLoad_Click, " & ex.Message & ", " & ex.TargetSite.Name)
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

    Private Sub btnDelError_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelError.Click
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
            savepage(True)
            PopPageTree()
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnDelError_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub BtnRawText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRawText.Click
        Try
            WebBrowser1.Document.Body.InnerHtml = txtRawText.Text
            Do While Not WebBrowser1.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            WebBrowser2.Document.Body.InnerHtml = txtRawText.Text
            Do While Not WebBrowser2.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            WebBrowser3.Document.Body.InnerHtml = txtRawText.Text
            Do While Not WebBrowser3.ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
            savepage(True)
            MblnDirty = False
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", BtnRawText_Click, " & ex.Message & ", " & ex.TargetSite.Name)
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

    Private Sub btnSplitPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSplitPage.Click
        Try
            Dim objXMLPage1 As MSXML2.IXMLDOMElement
            Dim DialogBox As New TextDialog()
            Dim blnDoit As Boolean
            Dim strPage As String

            blnDoit = True
            If MblnDirty Then
                Select Case MsgBox("Do you want to saved changes to the current page?" & vbCrLf & "Select Yes to save and create a new page, " & vbCrLf & "No lose changes or " & vbCrLf & "Cancel to to stay on this page", MsgBoxStyle.YesNoCancel, "Unsaved Changes")
                    Case MsgBoxResult.Yes
                        savepage(False)
                        MblnDirty = False
                    Case MsgBoxResult.Cancel
                        blnDoit = False
                End Select
            End If

            If blnDoit Then
                If DialogBox.ShowDialog = Windows.Forms.DialogResult.OK Then
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
                    MobjXMLPages.insertBefore(MobjXMLPage, objXMLPage1)
                    lblPage.Text = MstrPage
                    PopPageTree()
                    savepage(False)
                    TreeViewPages.SelectedNode = TreeViewPages.Nodes.Item(MstrPage)
                    MblnDirty = False
                    displaypage()
                End If
            End If
        Catch ex As Exception
            MobjLogWriter.WriteLine(Now().ToString("yyyy/MM/dd HH:mm:ss") & ", btnCopyPage_Click, " & ex.Message & ", " & ex.TargetSite.Name)
        End Try
    End Sub

    Private Sub btnSortPages_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSortPages.Click
        Dim blnFinished As Boolean
        Dim intLoop As Integer
        Dim objXMLPage1 As MSXML2.IXMLDOMElement
        Dim objXMLPage2 As MSXML2.IXMLDOMElement
        Dim strId1 As String
        Dim strId2 As String
        Me.Cursor = Cursors.WaitCursor
        Do
            blnFinished = True
            For intLoop = 0 To MobjXMLPages.childNodes.length - 2
                objXMLPage1 = MobjXMLPages.childNodes(intLoop)
                strId1 = getAttribute(objXMLPage1, "id")
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
            Next
            If blnFinished Then
                Exit Do
            End If
        Loop
        saveXml(MobjXMLDoc, TextBox1.Text)
        PopPageTree()
        TreeViewPages.SelectedNode = TreeViewPages.Nodes(0)
        Me.Cursor = Cursors.Default
    End Sub
End Class
