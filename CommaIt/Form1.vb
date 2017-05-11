Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnClear As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.txtInput = New System.Windows.Forms.TextBox()
        Me.gbAdd = New System.Windows.Forms.GroupBox()
        Me.rbNone = New System.Windows.Forms.RadioButton()
        Me.rbDouble = New System.Windows.Forms.RadioButton()
        Me.rbSingleQuote = New System.Windows.Forms.RadioButton()
        Me.btnDoit = New System.Windows.Forms.Button()
        Me.txtResult = New System.Windows.Forms.TextBox()
        Me.btnClearPaste = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbRpcNone = New System.Windows.Forms.RadioButton()
        Me.rbFirst = New System.Windows.Forms.RadioButton()
        Me.rbAll = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbFind = New System.Windows.Forms.TextBox()
        Me.tbReplace = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.gbAdd.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtInput
        '
        Me.txtInput.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtInput.Location = New System.Drawing.Point(8, 8)
        Me.txtInput.Multiline = True
        Me.txtInput.Name = "txtInput"
        Me.txtInput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInput.Size = New System.Drawing.Size(248, 432)
        Me.txtInput.TabIndex = 0
        Me.txtInput.WordWrap = False
        '
        'gbAdd
        '
        Me.gbAdd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbAdd.Controls.Add(Me.rbNone)
        Me.gbAdd.Controls.Add(Me.rbDouble)
        Me.gbAdd.Controls.Add(Me.rbSingleQuote)
        Me.gbAdd.Location = New System.Drawing.Point(264, 8)
        Me.gbAdd.Name = "gbAdd"
        Me.gbAdd.Size = New System.Drawing.Size(108, 96)
        Me.gbAdd.TabIndex = 1
        Me.gbAdd.TabStop = False
        Me.gbAdd.Text = "Add"
        '
        'rbNone
        '
        Me.rbNone.Location = New System.Drawing.Point(8, 64)
        Me.rbNone.Name = "rbNone"
        Me.rbNone.Size = New System.Drawing.Size(78, 24)
        Me.rbNone.TabIndex = 2
        Me.rbNone.Text = "None"
        '
        'rbDouble
        '
        Me.rbDouble.Location = New System.Drawing.Point(8, 40)
        Me.rbDouble.Name = "rbDouble"
        Me.rbDouble.Size = New System.Drawing.Size(94, 24)
        Me.rbDouble.TabIndex = 1
        Me.rbDouble.Text = "Double Quote"
        '
        'rbSingleQuote
        '
        Me.rbSingleQuote.Location = New System.Drawing.Point(8, 16)
        Me.rbSingleQuote.Name = "rbSingleQuote"
        Me.rbSingleQuote.Size = New System.Drawing.Size(94, 24)
        Me.rbSingleQuote.TabIndex = 0
        Me.rbSingleQuote.Text = "Single Quote"
        '
        'btnDoit
        '
        Me.btnDoit.Location = New System.Drawing.Point(264, 112)
        Me.btnDoit.Name = "btnDoit"
        Me.btnDoit.Size = New System.Drawing.Size(75, 23)
        Me.btnDoit.TabIndex = 2
        Me.btnDoit.Text = "Go"
        '
        'txtResult
        '
        Me.txtResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtResult.Location = New System.Drawing.Point(264, 144)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtResult.Size = New System.Drawing.Size(410, 336)
        Me.txtResult.TabIndex = 3
        Me.txtResult.WordWrap = False
        '
        'btnClearPaste
        '
        Me.btnClearPaste.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClearPaste.Location = New System.Drawing.Point(16, 448)
        Me.btnClearPaste.Name = "btnClearPaste"
        Me.btnClearPaste.Size = New System.Drawing.Size(104, 23)
        Me.btnClearPaste.TabIndex = 5
        Me.btnClearPaste.Text = "Clear, Paste && Go"
        '
        'btnClear
        '
        Me.btnClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClear.Location = New System.Drawing.Point(136, 448)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(104, 23)
        Me.btnClear.TabIndex = 6
        Me.btnClear.Text = "Clear"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.rbRpcNone)
        Me.GroupBox1.Controls.Add(Me.rbFirst)
        Me.GroupBox1.Controls.Add(Me.rbAll)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.tbFind)
        Me.GroupBox1.Controls.Add(Me.tbReplace)
        Me.GroupBox1.Location = New System.Drawing.Point(378, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(296, 96)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Replace"
        '
        'rbRpcNone
        '
        Me.rbRpcNone.AutoSize = True
        Me.rbRpcNone.Checked = True
        Me.rbRpcNone.Location = New System.Drawing.Point(14, 19)
        Me.rbRpcNone.Name = "rbRpcNone"
        Me.rbRpcNone.Size = New System.Drawing.Size(51, 17)
        Me.rbRpcNone.TabIndex = 6
        Me.rbRpcNone.TabStop = True
        Me.rbRpcNone.Text = "None"
        Me.rbRpcNone.UseVisualStyleBackColor = True
        '
        'rbFirst
        '
        Me.rbFirst.AutoSize = True
        Me.rbFirst.Location = New System.Drawing.Point(14, 64)
        Me.rbFirst.Name = "rbFirst"
        Me.rbFirst.Size = New System.Drawing.Size(88, 17)
        Me.rbFirst.TabIndex = 5
        Me.rbFirst.Text = "First Instance"
        Me.rbFirst.UseVisualStyleBackColor = True
        '
        'rbAll
        '
        Me.rbAll.AutoSize = True
        Me.rbAll.Location = New System.Drawing.Point(14, 41)
        Me.rbAll.Name = "rbAll"
        Me.rbAll.Size = New System.Drawing.Size(85, 17)
        Me.rbAll.TabIndex = 4
        Me.rbAll.Text = "All Instances"
        Me.rbAll.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(151, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Replace"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(171, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Find"
        '
        'tbFind
        '
        Me.tbFind.Location = New System.Drawing.Point(204, 26)
        Me.tbFind.Name = "tbFind"
        Me.tbFind.Size = New System.Drawing.Size(78, 20)
        Me.tbFind.TabIndex = 1
        '
        'tbReplace
        '
        Me.tbReplace.Location = New System.Drawing.Point(204, 52)
        Me.tbReplace.Name = "tbReplace"
        Me.tbReplace.Size = New System.Drawing.Size(78, 20)
        Me.tbReplace.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(582, 112)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(85, 23)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Save"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(679, 486)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnClearPaste)
        Me.Controls.Add(Me.txtResult)
        Me.Controls.Add(Me.btnDoit)
        Me.Controls.Add(Me.gbAdd)
        Me.Controls.Add(Me.txtInput)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Comma'It"
        Me.gbAdd.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub btnDoit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDoit.Click

        Copy()


    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.txtInput.Text = ""
    End Sub

    Private Sub btnClearPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearPaste.Click

        clearPaste()
        Copy()


    End Sub

    Private Sub Copy()
        Dim values As String() = Me.txtInput.Text.Replace(ControlChars.Lf, "").Split(ControlChars.Cr)

        Dim v As String
        Dim find As String
        Dim replace As String


        find = tbFind.Text.ToString
        replace = tbReplace.Text.ToString


        ''prefix selection
        Me.txtResult.Text = ""
        Dim prefixsuffix As String = ""

        If Me.rbDouble.Checked Then
            prefixsuffix = """"
        ElseIf Me.rbSingleQuote.Checked Then
            prefixsuffix = "'"
        End If

        Try
            For Each v In values
                v = v.Trim
                If rbAll.Checked = True Then

                    Dim v_replaced As String = v.Replace(find.ToString, replace.ToString)

                    v = v_replaced.ToString
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
                    Me.txtResult.Text += prefixsuffix & v & prefixsuffix & ","
                End If
            Next

            If Me.txtResult.Text.Substring(Me.txtResult.Text.Length - 1) = "," Then
                Me.txtResult.Text = Me.txtResult.Text.Substring(0, Me.txtResult.Text.Length - 1)
            End If

            Clipboard.SetDataObject(Me.txtResult.Text)
        Catch ex As Exception
            '  MessageBox.Show(ex.Message.ToString())
        End Try


    End Sub

    Private Sub clearPaste()
        'clear
        Me.txtInput.Text = ""
        'paste
        Dim iData As IDataObject = Clipboard.GetDataObject()
        ' Determines whether the data is in a format you can use.
        If iData.GetDataPresent(DataFormats.Text) Then
            ' Yes it is, so display it in a text box.
            Me.txtInput.Text = CType(iData.GetData(DataFormats.Text), String)
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK _
        Then
            My.Computer.FileSystem.WriteAllText _
            (SaveFileDialog1.FileName, txtInput.Text, True)
        End If
    End Sub
End Class
