Attribute VB_Name = "Module1"
Option Explicit
'64bit����p�̂�
'�L���v�`���[
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'�g�[�X�g
Private Declare PtrSafe Function MessageBoxTimeoutA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As VbMsgBoxStyle, ByVal wLanguageID As Long, ByVal dwMilliseconds As Long) As Long
'���j���[
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
'�萔
Private Const GWL_STYLE = -16
Private Const SC_CLOSE = &HF060      '���j���[�́u�~�v�i����j
Private Const MF_BYCOMMAND = &H0&    '�萔�̐ݒ�
Private Const WS_THICKFRAME = &H40000 '�E�B���h�E�̃T�C�Y�ύX
Private Const WS_MINIMIZEBOX = &H20000 '�ŏ����{�^��
Private Const WS_MAXIMIZEBOX = &H10000 '�ő剻�{�^��
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
'�\����
Type RGBNumber
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

'�I���t���O
Dim bExitFlg As Boolean

'�L���v�`���[���C������
Sub CaptureCore(pWb As Workbook)
    OpenClipboard 0&
    EmptyClipboard
    CloseClipboard
    Dim CB As Variant
    Dim position As Integer: position = 10
    Dim size As Double
    Dim typRBG As RGBNumber
    ' �I�����Ă���Z������Z���Ƃ��Ď擾����
    Dim baseCell As Variant
    Dim i As Integer
    Dim objShp As Variant
    Do While True
        CB = Application.ClipboardFormats
        If bExitFlg = True Then GoTo Quit

        On Error GoTo ErrorQuit
        For i = 1 To UBound(CB)
            If CB(i) = xlClipboardFormatBitmap Then
                With UserForm1
                    .lblStatus.Caption = "****"
                    .Repaint
                    size = CInt(.cmbRate.Value) / 100
                    typRBG = ColorToRGB(.lblColor.BackColor)
                End With

                pWb.Activate
                Set baseCell = Selection
                If UserForm1.chkTimeOutput.Value Then
                    ' ����������
                    Range(baseCell.Offset(0, 0), baseCell.Offset(0, 1 + 10)).Select
                    With Selection.Borders(xlEdgeBottom)
                      .LineStyle = xlContinuous
                      .Weight = xlThin
                    End With

                    ' ���o���p�̋L�����Z�b�g����
                    baseCell.Offset(0, 0).Value = "��"
                      ' �L���v�`���擾�������Z�b�g����
                    With baseCell.Offset(0, 1)
                      .HorizontalAlignment = xlLeft
                      .Value = "�擾�����F" & Now
                    End With
                    ActiveCell.Offset(1, 1).Select
                End If
    
                Sleep 1000

                ActiveSheet.Paste

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
                DoEvents
                Call ToastMsg("�L���v�`���[����")
                UserForm1.Repaint
            End If
        Next i
        DoEvents
    Loop

Quit:
    Call ToastMsg("�I�����܂��B")
    Call SaveWorkBook(pWb)
    pWb.Close
    GoTo ToEnd
ErrorQuit:
    MsgBox "�\�����ʓ���̂��ߒ�~���܂����B", vbInformation
ToEnd:
End Sub

'�@�t�H�[�������J�n
Public Sub Begin()
    Dim wRet As Long
    Dim wStyle As Long
    Dim hMenu As LongPtr
    Dim rClose As Long
    Dim strFileName As String
    Dim hwnd As LongPtr
    strFileName = UserForm1.txtCustomName.Text + "_" + Format(Date, "yyyymmdd") + ".xlsx"
    With UserForm1
        .tglStart.Enabled = False
        .tglStart.Caption = "��~��"
        .chkTimeOutput.Value = True
        .Caption = "Start Capturing"
        .lblPath.Caption = ThisWorkbook.Path
        .lblFile.Caption = strFileName
        .lblColor.BackColor = RGB(0, 0, 0)
        .Show vbModeless
    End With
    hwnd = SetMenuButton
'    UserForm1.Repaint
    Call pouseCold

    Windows(ThisWorkbook.Name).WindowState = xlMinimized
    '�ŏ�������ꍇ
    Call ShowWindow(hwnd, SW_SHOWMINIMIZED)
    DoEvents
    ' Excel���ŏ���
    Application.WindowState = xlMinimized
    ' �L���v�`���[�����J�n
    Call StartCapture(strFileName)

    Exit Sub
End Sub

' �t�H�[�����j���[����
Private Function SetMenuButton() As LongPtr
    Dim hwnd As LongPtr
    Dim wRet As Long
    Dim wStyle As Long
    Dim hMenu As LongPtr
    Dim rClose As Long
    
    hwnd = FindWindow("ThunderDFrame", UserForm1.Caption) 'Window�n���h���擾
    wStyle = GetWindowLong(hwnd, GWL_STYLE)  'Use��Form��Window���擾
    'wStyle = (wStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)  'Min,Max���j���[�{�^���A�T�C�Y�ύX��ǉ�
    wStyle = (wStyle Or WS_THICKFRAME Or WS_MINIMIZEBOX)  'Min,Max���j���[�{�^���A�T�C�Y�ύX��ǉ�
    wRet = SetWindowLong(hwnd, GWL_STYLE, wStyle) '�ǉ������{�^���ݒ�
    hMenu = GetSystemMenu(hwnd, 0&) '���j���[���擾
    rClose = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)  '�u�~�v�{�^���̖�����
    wRet = DrawMenuBar(hwnd) 'UseForm�̃��j���[�o�[�O�g���ĕ`��
    UserForm1.Repaint
    
    SetMenuButton = hwnd
End Function

' �L���v�`���[�����J�n
Public Sub StartCapture(pStrFileName As String)
    Dim wb As Workbook
    bExitFlg = False
    ' �o�̓u�b�N����
    Set wb = SetNewBook(pStrFileName)
    Call pouseCold
    With UserForm1
        .tglStart.Enabled = True
        .tglStart.Caption = "�N����"
    End With

    Windows(pStrFileName).WindowState = xlMinimized
    '�L���v�`���[���C������
    Call CaptureCore(wb)

End Sub

' �o�̓u�b�N����
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

' �t�@�C���ۑ�����
Private Sub SaveWorkBook(pWb As Workbook, Optional pStrFullFileName As String)
    '�x�����b�Z�[�W��\�����Ȃ�
    Application.DisplayAlerts = False
    If Len(pStrFullFileName) > 0 Then
        pWb.SaveAs pStrFullFileName
    Else
        pWb.Save
    End If
    '�x�����b�Z�[�W��\������
    Application.DisplayAlerts = True
'    Windows(pStrFileName).WindowState = xlMinimized
End Sub

' �o�̓V�[�g���擾����
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

' �V�[�g���݃`�F�b�N
Private Function IsSheetExists(pSheetName As String) As Boolean
    Dim ws As Worksheet, flag As Boolean
    flag = False
    For Each ws In Worksheets
        If ws.Name = pSheetName Then
            flag = True
            Exit For
        End If
    Next ws
    IsSheetExists = flag
End Function


' �����L���v�`���[���I������
Public Sub StopCapture()
    bExitFlg = True
End Sub

' �`��p�ɃX���[�v����
Public Sub pouseCold()
    DoEvents
    Sleep (1000)
End Sub

' �g�[�X�g�\��
Sub ToastMsg(strMsg As String)
    Dim rtn As Long
    Dim sTitle As String
    Dim sMsg As String
    Dim intMilliSecound As Integer
    Call pouseCold

    sTitle = "�ʒm"
    sMsg = "���b�Z�[�W"
    intMilliSecound = 1000
    'API
    rtn = MessageBoxTimeoutA(0&, strMsg, sTitle, vbOKOnly, 0&, intMilliSecound)
End Sub

' COLOR�R�[�h��RGB�p�\���̂ɕϊ�����
Private Function ColorToRGB(ByVal lngColor As Long) As RGBNumber
    With ColorToRGB
        .Red = lngColor Mod 256
        .Green = Int(lngColor / 256) Mod 256
        .Blue = Int(lngColor / 256 / 256)
    End With
End Function
