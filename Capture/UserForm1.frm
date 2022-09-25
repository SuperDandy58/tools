VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Start Capturing"
   ClientHeight    =   2202
   ClientLeft      =   -1176
   ClientTop       =   -5640
   ClientWidth     =   5934
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Private Sub cmdFileDialog_Click()
'    Dim PathName As String, FileName As String
'    With Application.FileDialog(msoFileDialogFilePicker)
'        .Filters.Clear
'        .Filters.Add "Excelファイル", "*.xlsx"
'        .InitialFileName = ThisWorkbook.Path & "\"
'        .AllowMultiSelect = False
'        If .Show = True Then
'            txtCustomName.Text = .SelectedItems(1)
'
'
'            FileName = Dir(.SelectedItems(1))
'            PathName = Replace(.SelectedItems(1), FileName, "")
'            txtCustomName.Text = FileName
'        End If
'    End With
'End Sub

Private Sub UserForm_Initialize()
  Dim strRateValue As Variant
  Dim rateValue As String
  Dim iCnt As Integer
  strRateValue = Array("10", "20", "30", "40", "50", "60", "70", "80", "90", "100")
 
  For iCnt = 0 To UBound(strRateValue)
       cmbRate.AddItem strRateValue(iCnt)
  Next
   
  cmbRate.ListIndex = 1
End Sub


Private Sub cmdClose_Click()
    Call StopCapture
    ' 画面閉じる
    Unload UserForm1
End Sub

Private Sub cmdColorSelect_Click()
    Application.Dialogs(xlDialogEditColor).Show (1)
    UserForm1.lblColor.BackColor = ActiveWorkbook.Colors(1)
End Sub

Private Sub cmdReStart_Click()
    Dim strFileName As String
    tglStart.Caption = "停止中"
    tglStart.Enabled = False
    
    Me.Repaint
     Call pouseCold
    tglStart.Caption = "起動中"
    tglStart.Enabled = True
    
    Call StopCapture

    strFileName = UserForm1.txtCustomName.Text + "_" + Format(Date, "yyyymmdd") + ".xlsx"
    lblFile.Caption = strFileName
    Call StartCapture(strFileName)
    
End Sub




