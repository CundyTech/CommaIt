Public Class MainForm
    Inherits Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtInput As System.Windows.Forms.TextBox
    Friend WithEvents gbAdd As System.Windows.Forms.GroupBox
    Friend WithEvents rbSingleQuote As System.Windows.Forms.RadioButton
    Friend WithEvents btnDoit As System.Windows.Forms.Button
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents rbDouble As System.Windows.Forms.RadioButton
    Friend WithEvents rbNone As System.Windows.Forms.RadioButton
    Friend WithEvents btnClearPaste As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbFirst As System.Windows.Forms.RadioButton
    Friend WithEvents rbAll As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbFind As System.Windows.Forms.TextBox
    Friend WithEvents tbReplace As System.Windows.Forms.TextBox
    Friend WithEvents rbRpcNone As System.Windows.Forms.RadioButton
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnClear As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        txtInput = New System.Windows.Forms.TextBox()
        gbAdd = New System.Windows.Forms.GroupBox()
        rbNone = New System.Windows.Forms.RadioButton()
        rbDouble = New System.Windows.Forms.RadioButton()
        rbSingleQuote = New System.Windows.Forms.RadioButton()
        btnDoit = New System.Windows.Forms.Button()
        txtResult = New System.Windows.Forms.TextBox()
        btnClearPaste = New System.Windows.Forms.Button()
        btnClear = New System.Windows.Forms.Button()
        GroupBox1 = New System.Windows.Forms.GroupBox()
        rbRpcNone = New System.Windows.Forms.RadioButton()
        rbFirst = New System.Windows.Forms.RadioButton()
        rbAll = New System.Windows.Forms.RadioButton()
        Label2 = New System.Windows.Forms.Label()
        Label1 = New System.Windows.Forms.Label()
        tbFind = New System.Windows.Forms.TextBox()
        tbReplace = New System.Windows.Forms.TextBox()
        btnSave = New System.Windows.Forms.Button()
        SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        gbAdd.SuspendLayout()
        GroupBox1.SuspendLayout()
        SuspendLayout()
        '
        'txtInput
        '
        txtInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        txtInput.Location = New System.Drawing.Point(8, 8)
        txtInput.Multiline = True
        txtInput.Name = "txtInput"
        txtInput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txtInput.Size = New System.Drawing.Size(248, 432)
        txtInput.TabIndex = 0
        txtInput.WordWrap = False
        '
        'gbAdd
        '
        gbAdd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        gbAdd.Controls.Add(rbNone)
        gbAdd.Controls.Add(rbDouble)
        gbAdd.Controls.Add(rbSingleQuote)
        gbAdd.Location = New System.Drawing.Point(264, 8)
        gbAdd.Name = "gbAdd"
        gbAdd.Size = New System.Drawing.Size(108, 96)
        gbAdd.TabIndex = 1
        gbAdd.TabStop = False
        gbAdd.Text = "Add"
        '
        'rbNone
        '
        rbNone.Location = New System.Drawing.Point(8, 64)
        rbNone.Name = "rbNone"
        rbNone.Size = New System.Drawing.Size(78, 24)
        rbNone.TabIndex = 2
        rbNone.Text = "None"
        '
        'rbDouble
        '
        rbDouble.Location = New System.Drawing.Point(8, 40)
        rbDouble.Name = "rbDouble"
        rbDouble.Size = New System.Drawing.Size(94, 24)
        rbDouble.TabIndex = 1
        rbDouble.Text = "Double Quote"
        '
        'rbSingleQuote
        '
        rbSingleQuote.Location = New System.Drawing.Point(8, 16)
        rbSingleQuote.Name = "rbSingleQuote"
        rbSingleQuote.Size = New System.Drawing.Size(94, 24)
        rbSingleQuote.TabIndex = 0
        rbSingleQuote.Text = "Single Quote"
        '
        'btnDoit
        '
        btnDoit.Location = New System.Drawing.Point(264, 112)
        btnDoit.Name = "btnDoit"
        btnDoit.Size = New System.Drawing.Size(75, 23)
        btnDoit.TabIndex = 2
        btnDoit.Text = "Go"
        '
        'txtResult
        '
        txtResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        txtResult.Location = New System.Drawing.Point(264, 144)
        txtResult.Multiline = True
        txtResult.Name = "txtResult"
        txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txtResult.Size = New System.Drawing.Size(410, 336)
        txtResult.TabIndex = 3
        txtResult.WordWrap = False
        '
        'btnClearPaste
        '
        btnClearPaste.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        btnClearPaste.Location = New System.Drawing.Point(16, 448)
        btnClearPaste.Name = "btnClearPaste"
        btnClearPaste.Size = New System.Drawing.Size(104, 23)
        btnClearPaste.TabIndex = 5
        btnClearPaste.Text = "Clear, Paste && Go"
        '
        'btnClear
        '
        btnClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        btnClear.Location = New System.Drawing.Point(136, 448)
        btnClear.Name = "btnClear"
        btnClear.Size = New System.Drawing.Size(104, 23)
        btnClear.TabIndex = 6
        btnClear.Text = "Clear"
        '
        'GroupBox1
        '
        GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        GroupBox1.Controls.Add(rbRpcNone)
        GroupBox1.Controls.Add(rbFirst)
        GroupBox1.Controls.Add(rbAll)
        GroupBox1.Controls.Add(Label2)
        GroupBox1.Controls.Add(Label1)
        GroupBox1.Controls.Add(tbFind)
        GroupBox1.Controls.Add(tbReplace)
        GroupBox1.Location = New System.Drawing.Point(378, 8)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New System.Drawing.Size(296, 96)
        GroupBox1.TabIndex = 7
        GroupBox1.TabStop = False
        GroupBox1.Text = "Replace"
        '
        'rbRpcNone
        '
        rbRpcNone.AutoSize = True
        rbRpcNone.Checked = True
        rbRpcNone.Location = New System.Drawing.Point(14, 19)
        rbRpcNone.Name = "rbRpcNone"
        rbRpcNone.Size = New System.Drawing.Size(51, 17)
        rbRpcNone.TabIndex = 6
        rbRpcNone.TabStop = True
        rbRpcNone.Text = "None"
        rbRpcNone.UseVisualStyleBackColor = True
        '
        'rbFirst
        '
        rbFirst.AutoSize = True
        rbFirst.Location = New System.Drawing.Point(14, 64)
        rbFirst.Name = "rbFirst"
        rbFirst.Size = New System.Drawing.Size(88, 17)
        rbFirst.TabIndex = 5
        rbFirst.Text = "First Instance"
        rbFirst.UseVisualStyleBackColor = True
        '
        'rbAll
        '
        rbAll.AutoSize = True
        rbAll.Location = New System.Drawing.Point(14, 41)
        rbAll.Name = "rbAll"
        rbAll.Size = New System.Drawing.Size(85, 17)
        rbAll.TabIndex = 4
        rbAll.Text = "All Instances"
        rbAll.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(151, 59)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(47, 13)
        Label2.TabIndex = 3
        Label2.Text = "Replace"
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(171, 33)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(27, 13)
        Label1.TabIndex = 2
        Label1.Text = "Find"
        '
        'tbFind
        '
        tbFind.Location = New System.Drawing.Point(204, 26)
        tbFind.Name = "tbFind"
        tbFind.Size = New System.Drawing.Size(78, 20)
        tbFind.TabIndex = 1
        '
        'tbReplace
        '
        tbReplace.Location = New System.Drawing.Point(204, 52)
        tbReplace.Name = "tbReplace"
        tbReplace.Size = New System.Drawing.Size(78, 20)
        tbReplace.TabIndex = 0
        '
        'btnSave
        '
        btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        btnSave.Location = New System.Drawing.Point(582, 112)
        btnSave.Name = "btnSave"
        btnSave.Size = New System.Drawing.Size(85, 23)
        btnSave.TabIndex = 8
        btnSave.Text = "Save"
        '
        'MainForm
        '
        AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        ClientSize = New System.Drawing.Size(679, 486)
        Controls.Add(btnSave)
        Controls.Add(GroupBox1)
        Controls.Add(btnClear)
        Controls.Add(btnClearPaste)
        Controls.Add(txtResult)
        Controls.Add(btnDoit)
        Controls.Add(gbAdd)
        Controls.Add(txtInput)
        Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Name = "MainForm"
        Text = "Comma'It"
        gbAdd.ResumeLayout(False)
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()

    End Sub

#End Region

    ''' <summary>
    ''' Go btn click event.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnDoit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDoit.Click

        CommaSeparate()
        
    End Sub

    ''' <summary>
    ''' Clear text input textbox.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClear.Click

        txtInput.Text = ""

    End Sub

    ''' <summary>
    ''' Clear text input, paste clipboard, then comma separate the list and then CommaSeparate results back to clipboard.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnClearPaste_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClearPaste.Click

        ClearPaste()
        CommaSeparate()

    End Sub

    ''' <summary>
    ''' Comma separate pasted text, 
    ''' </summary>
    Private Sub CommaSeparate()

        Dim values As String() = txtInput.Text.Replace(ControlChars.Lf, "").Split(ControlChars.Cr)
        Dim v As String
        Dim find As String
        Dim replace As String

        find = tbFind.Text.ToString
        replace = tbReplace.Text.ToString

        ''Prefix selection
        txtResult.Text = ""
        Dim prefixsuffix As String = ""

        If rbDouble.Checked Then

            prefixsuffix = """"

        ElseIf rbSingleQuote.Checked Then

            prefixsuffix = "'"

        End If

        ''Do any find and replace and then comma separate
        Try
            For Each v In values

                v = v.Trim
                If rbAll.Checked = True Then

                    Dim vReplaced As String = v.Replace(find.ToString, replace.ToString)
                    v = vReplaced.ToString

                ElseIf rbFirst.Checked = True Then

                    Dim i As String = v.IndexOf(find)

                    If i = -1 Then

                        v = v

                    Else

                        Dim newString As String = v.Substring(0, i) & replace & v.Substring(i + find.Length)
                        v = newString

                    End If

                End If

                If v.Length > 0 Then

                    txtResult.Text += prefixsuffix & v & prefixsuffix & ","

                End If
            Next

            If txtResult.Text.Substring(txtResult.Text.Length - 1) = "," Then

                txtResult.Text = txtResult.Text.Substring(0, txtResult.Text.Length - 1)

            End If

            Clipboard.SetDataObject(txtResult.Text)

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' Clear text input and paste clipboard
    ''' </summary>
    Private Sub ClearPaste()

        'clear
        txtInput.Text = ""

        'paste
        Dim iData As IDataObject = Clipboard.GetDataObject()

        ' Determines whether the data is in a format you can use.
        If iData.GetDataPresent(DataFormats.Text) Then

            ' Yes it is, so display it in a text box.
            txtInput.Text = CType(iData.GetData(DataFormats.Text), String)

        End If

    End Sub

    ''' <summary>
    ''' Save output to a text file
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

            My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, txtInput.Text, True)

        End If

    End Sub

End Class
