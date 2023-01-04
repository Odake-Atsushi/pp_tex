Attribute VB_Name = "TexModule"
Option Explicit
' =====================================================
Const ControlCaption As String = "TexInput"

Public Sub Auto_Open()
'アドインとして読み込んだときに実行
  With Application.CommandBars("Standard")
    On Error Resume Next
    .Controls(ControlCaption).Delete
    On Error GoTo 0
    With .Controls.Add(Type:=msoControlButton) '[アドイン]タブにボタン追加
      .Caption = ControlCaption
      .Style = msoButtonIconAndCaption
      .FaceId = 65 'アイコン画像の設定
      .OnAction = "Add_Tex"
    End With
  End With
End Sub

Public Sub Auto_Close()
'アドインの読み込み解除したときに実行
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
    ' texコンパイル
    Dim execCommand As String
    Dim wsh As Object
    Dim result As Integer
    Dim folder_path As String
    folder_path = "C:\Windows\Temp\TexInput"
    '実行するコマンドを指定
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = folder_path
    
    'コマンドを同期実行
    execCommand = "pdfcrop texinput_buf.pdf"
    result = wsh.run(Command:="%ComSpec% /c " & execCommand, WindowStyle:=0, WaitOnReturn:=True)
    
    If (result <> 0) Then
        MsgBox ("pdfcrop コマンドが異常終了しました。")
        '後片付け
        Set wsh = Nothing
    Else
        'コマンドを同期実行
        execCommand = "dvisvgm --pdf texinput_buf-crop.pdf"
        result = wsh.run(Command:="%ComSpec% /c " & execCommand, WindowStyle:=0, WaitOnReturn:=True)
    
        If (result <> 0) Then
            MsgBox ("dvisvgm コマンドが異常終了しました。")
        End If
    End If
    
    '後片付け
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


