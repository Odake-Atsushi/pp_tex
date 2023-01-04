Attribute VB_Name = "TexModule"
Option Explicit
' =====================================================
Const ControlCaption As String = "TexInput"

Public Sub Auto_Open()
'�A�h�C���Ƃ��ēǂݍ��񂾂Ƃ��Ɏ��s
  With Application.CommandBars("Standard")
    On Error Resume Next
    .Controls(ControlCaption).Delete
    On Error GoTo 0
    With .Controls.Add(Type:=msoControlButton) '[�A�h�C��]�^�u�Ƀ{�^���ǉ�
      .Caption = ControlCaption
      .Style = msoButtonIconAndCaption
      .FaceId = 65 '�A�C�R���摜�̐ݒ�
      .OnAction = "Add_Tex"
    End With
  End With
End Sub

Public Sub Auto_Close()
'�A�h�C���̓ǂݍ��݉��������Ƃ��Ɏ��s
  On Error Resume Next
  Application.CommandBars("Standard").Controls(ControlCaption).Delete
  On Error GoTo 0
End Sub
' =====================================================


Sub Add_Tex()
    UserForm1.Show vbModeless
End Sub

Function pdf_to_svg()
    ' =====================================================
    ' tex�R���p�C��
    Dim execCommand As String
    Dim wsh As Object
    Dim result As Integer
    Dim folder_path As String
    folder_path = "C:\Windows\Temp\TexInput"
    '���s����R�}���h���w��
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = folder_path
    
    '�R�}���h�𓯊����s
    execCommand = "pdfcrop texinput_buf.pdf"
    result = wsh.run(Command:="%ComSpec% /c " & execCommand, WindowStyle:=0, WaitOnReturn:=True)
    
    If (result <> 0) Then
        MsgBox ("pdfcrop �R�}���h���ُ�I�����܂����B")
        '��Еt��
        Set wsh = Nothing
    Else
        '�R�}���h�𓯊����s
        execCommand = "dvisvgm --pdf texinput_buf-crop.pdf"
        result = wsh.run(Command:="%ComSpec% /c " & execCommand, WindowStyle:=0, WaitOnReturn:=True)
    
        If (result <> 0) Then
            MsgBox ("dvisvgm �R�}���h���ُ�I�����܂����B")
        End If
    End If
    
    '��Еt��
    Set wsh = Nothing
    ' =====================================================
    
    Dim svg_path As String
    svg_path = folder_path & "\texinput_buf-crop.svg"
    Call add_picture(svg_path)
End Function
Function add_picture(ByVal picture_path As String)
    Dim myDocument As Object
    Set myDocument = ActivePresentation.Slides(ActiveWindow.Selection.SlideRange.SlideIndex)
    myDocument.Shapes.AddPicture _
        FileName:=picture_path, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=100, _
        Top:=100
End Function


