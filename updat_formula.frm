VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   OleObjectBlob   =   "updat_formula.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Add reference: Microsoft VBscript Regular Expression

Private Sub CommandButton1_Click()
   
Dim fs, f, f1, fc, s, x, rowss, columnss

  '** 使用FileDialog对象来选择文件夹
Dim fd As FileDialog
Dim strPath As String
       
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
       
        '** 显示选择文件夹对话框
If fd.Show = -1 Then        '** 用户选择了文件夹
    f_path = fd.SelectedItems(1)
Else
    f_path = ""
End If
    Set fd = Nothing

Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(f_path) 'Directory of excel files will be merge
Set fc = f.Files

For Each f1 In fc

ListBox1.AddItem (f1.Path)

Next

End Sub

Private Sub CommandButton2_Click()

Set excelObj = CreateObject("Excel.Application")

excelObj.Visible = True

Set myExcel = Application.Workbooks.Open(ListBox1.Text)

Set myrange = Application.InputBox(prompt:="选择单元格：", Type:=8)

Dim formular_str As String
formular_str = myrange.Formula
TextBox1.Text = formular_str

Set regx_1 = CreateObject("VBscript.regexp")
regx_1.Pattern = "[A-Z][A-Z]?[0-9]+"
regx_1.Global = True

Set regx_2 = CreateObject("VBscript.regexp")
regx_2.Pattern = "[0-9]"
regx_2.Global = True

Set matches = regx_1.Execute(formular_str)

For Each mtch In matches

    If ListBox2.ListCount = 0 Then
        
        ListBox2.AddItem (mtch)
    
        ListBox4.AddItem (Range(regx_2.Replace(mtch, "1")).Value)
    
    Else

        j = 0
    
        For i = 0 To ListBox2.ListCount - 1
        
            If mtch = ListBox2.List(i) Then
                j = j + 1
            End If
            
        Next i
        
        If j = 0 Then
        
            ListBox2.AddItem (mtch)
        
            ListBox4.AddItem (Range(regx_2.Replace(mtch, "1")).Value)
        
        End If
    
    End If
Next

Set myrange = Application.InputBox(prompt:="选择单元格范围：", Type:=8)

Dim str As String
For i = 0 To ListBox2.ListCount - 1

    For Each myrng In myrange
    
        If myrng.Value = ListBox4.List(i) Then
        
            str = Replace(myrng.Address, "$", "")
            ListBox3.AddItem (Replace(str, "1", "2"))
            
        End If
    Next

Next i


'myExcel.Close
'
'excelObj.Quit
'
'Set excelObj = Nothing

End Sub

Private Sub CommandButton3_Click()


TextBox2.Text = TextBox1.Text

For i = 0 To ListBox3.ListCount - 1

    MsgBox ListBox3.List(i)
    
    TextBox2.Text = Replace(TextBox2.Text, ListBox2.List(i), ListBox3.List(i))

Next i

Set myrange = Application.InputBox(prompt:="选择原公式单元格", Type:=8)

myrange.Formula = TextBox2.Text

End Sub

