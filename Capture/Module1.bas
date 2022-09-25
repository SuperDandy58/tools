Attribute VB_Name = "Module1"
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function MessageBoxTimeoutA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As VbMsgBoxStyle, ByVal wLanguageID As Long, ByVal dwMilliseconds As Long) As Long


'64bit動作用
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Const GWL_STYLE = -16
Private Const SC_CLOSE = &HF060      'メニューの「×」（閉じる）
Private Const MF_BYCOMMAND = &H0&    '定数の設定
Private Const WS_THICKFRAME = &H40000 'ウィンドウのサイズ変更
Private Const WS_MINIMIZEBOX = &H20000 '最小化ボタン
Private Const WS_MAXIMIZEBOX = &H10000 '最大化ボタン
 
Type RGBNumber
    Red As Byte
    Green As Byte
    Blue As Byte
End Type
Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
'Private Declare Function ShowWindow Lib "User32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
'
Dim hwnd As LongPtr

Dim bExitFlg As Boolean

Sub CaptureCore(pWb As Workbook)
    OpenClipboard 0&
    EmptyClipboard
    CloseClipboard
    Dim CB As Variant
    Dim position As Integer: position = 10
    Dim size As Double
    
    ' 選択しているセルを基準セルとして取得する
    Dim baseCell As Variant

Dim typRBG As RGBNumber
    Do While True
        CB = Application.ClipboardFormats
        If bExitFlg = True Then GoTo Quit

'        Dim rows As Integer: rows = 54   ' 行数
'        Dim cols As Integer: cols = 72   ' 列数
        On Error GoTo ErrorQuit
        For i = 1 To UBound(CB)
            If CB(i) = xlClipboardFormatBitmap Then
                UserForm1.lblStatus.Caption = "****"
                UserForm1.Repaint
                size = CInt(UserForm1.cmbRate.Value) / 100
                pWb.Activate
                Set baseCell = Selection
                ' 下線を引く
                Range(baseCell.Offset(0, 0), baseCell.Offset(0, 1 + 10)).Select
                With Selection.Borders(xlEdgeBottom)
                  .LineStyle = xlContinuous
                  .Weight = xlThin
                End With

                ' 見出し用の記号をセットする
                baseCell.Offset(0, 0).Value = "■"
                  ' キャプチャ取得日時をセットする
                With baseCell.Offset(0, 1)
                  .HorizontalAlignment = xlLeft
                  .Value = "取得日時：" & Now
                End With
                ActiveCell.Offset(1, 1).Select

    
                Sleep 1000
                
                ActiveSheet.Paste
                typRBG = Color_to_RGB(UserForm1.lblColor.BackColor)
                Set objShp = ActiveSheet.Shapes(Selection.Name)
                With objShp
                    .LockAspectRatio = msoTrue
                    .ScaleHeight size, msoTrue
                    .Line.Visible = msoTrue
                    .Line.ForeColor.RGB = RGB(typRBG.Red, typRBG.Green, typRBG.Blue)
                    .Line.Weight = 1
                    Cells(.BottomRightCell.Row + 2, .TopLeftCell.Column - 1).Select
                End With
                
                OpenClipboard 0&
                EmptyClipboard
                CloseClipboard
                UserForm1.lblStatus.Caption = "OK"
                Call SaveWorkBook(pWb)
                Call ToastMsg("キャプチャー成功")
                UserForm1.Repaint
            End If
        Next i
        DoEvents
    Loop

Quit:
    Call ToastMsg("終了します。")
    Call SaveWorkBook(pWb)
    pWb.Close
    GoTo ToEnd
ErrorQuit:
    MsgBox "予期せぬ動作のため停止しました。", vbInformation
ToEnd:
End Sub


Public Sub Begin()
    Dim wRet As Long
    Dim wStyle As Long
    Dim hMenu As LongPtr
    Dim rClose As Long
    Dim strFileName As String
    
    strFileName = UserForm1.txtCustomName.Text + "_" + Format(Date, "yyyymmdd") + ".xlsx"
    With UserForm1
        .tglStart.Enabled = False
        .tglStart.Caption = "停止中"
        .chkTimeOut.Value = True
    
        .Caption = "Start Capturing"
        .lblPath.Caption = ThisWorkbook.Path
        .lblFile.Caption = strFileName
        .lblColor.BackColor = RGB(0, 0, 0)
        .Show vbModelesss
    End With
    Call SetMenuButton
'    UserForm1.Repaint
    Call pouseCold

    
    Windows(ThisWorkbook.Name).WindowState = xlMinimized
    '最小化する場合
    Call ShowWindow(hwnd, SW_SHOWMINIMIZED)
    DoEvents
    ' Excelを最小化
    Application.WindowState = xlMinimized

    Call StartCapture(strFileName)

    Exit Sub
End Sub

Private Sub SetMenuButton()
    hwnd = FindWindow("ThunderDFrame", UserForm1.Caption) 'Windowハンドル取得
    wStyle = GetWindowLong(hwnd, GWL_STYLE)  'UseｒFormのWindow情報取得
    'wStyle = (wStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)  'Min,Maxメニューボタン、サイズ変更を追加
    wStyle = (wStyle Or WS_THICKFRAME Or WS_MINIMIZEBOX)  'Min,Maxメニューボタン、サイズ変更を追加
    wRet = SetWindowLong(hwnd, GWL_STYLE, wStyle) '追加したボタン設定
    hMenu = GetSystemMenu(hwnd, 0&) 'メニュー情報取得
    rClose = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)  '「×」ボタンの無効化
    wRet = DrawMenuBar(hwnd) 'UseFormのメニューバー外枠を再描画
    UserForm1.Repaint

End Sub



Public Sub StartCapture(pStrFileName As String)
    Dim wb As Workbook
    bExitFlg = False
    Set wb = SetNewBook(pStrFileName)
    Call pouseCold
    With UserForm1
        .tglStart.Enabled = True
        .tglStart.Caption = "起動中"
    End With

    Windows(pStrFileName).WindowState = xlMinimized
    Call CaptureCore(wb)

End Sub

Private Function SetNewBook(pStrFileName As String) As Workbook
    Dim NewBook As Workbook
    Dim strFullFileName As String
 
    strFullFileName = ThisWorkbook.Path & "\" & pStrFileName

    If Dir(strFullFileName) = "" Then
        Set NewBook = Workbooks.Add
    Else
        Set NewBook = Workbooks.Open(strFullFileName)
    End If
    Worksheets.Add
    ActiveSheet.Name = getNewSheetName()
    Call SaveWorkBook(NewBook, strFullFileName)

    NewBook.Activate
    DoEvents

    Set SetNewBook = NewBook
End Function


Private Sub SaveWorkBook(pWb As Workbook, Optional pStrFullFileName As String)
    '警告メッセージを表示しない
    Application.DisplayAlerts = False
    If Len(pStrFullFileName) > 0 Then
        pWb.SaveAs pStrFullFileName
    Else
        pWb.Save
    End If
    '警告メッセージを表示する
    Application.DisplayAlerts = True
'    Windows(pStrFileName).WindowState = xlMinimized
End Sub


Private Function getNewSheetName() As String
    Dim strSheetName As String
    Dim iCount As Integer
    iCount = 1
    strSheetName = "CP" & Format(iCount, "000")
    Do
        If Not IsSheetExists(strSheetName) Then
            Exit Do
        End If
        iCount = iCount + 1
        strSheetName = "CP" & Format(iCount, "000")
    Loop
    getNewSheetName = strSheetName
End Function

Private Function IsSheetExists(pSheetName As String) As Boolean
    Dim ws As Worksheet, flag As Boolean
    flag = Fals
    For Each ws In Worksheets
        If ws.Name = pSheetName Then
            flag = True
            Exit For
        End If
    Next ws
    IsSheetExists = flag
End Function



Public Sub StopCapture()
    bExitFlg = True
End Sub


Public Sub pouseCold()
    DoEvents
    Sleep (1000)
End Sub


Sub ToastMsg(strMsg As String)
    Dim rtn As Long
    Dim sTitle As String
    Dim sMsg As String
    Dim intMilliSecound As Integer
         
    sTitle = "通知"
    sMsg = "メッセージ"
    intMilliSecound = 1000
    'API
    rtn = MessageBoxTimeoutA(0&, strMsg, sTitle, vbOKOnly, 0&, intMilliSecound)
End Sub


Private Function Color_to_RGB(ByVal lngColor As Long) As RGBNumber
    With Color_to_RGB
        .Red = lngColor Mod 256
        .Green = Int(lngColor / 256) Mod 256
        .Blue = Int(lngColor / 256 / 256)
    End With
End Function
