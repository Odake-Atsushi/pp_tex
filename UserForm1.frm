VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6590
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7060
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    ' フォルダ作成
    Dim folder_path As String
    folder_path = "C:\Windows\Temp\TexInput"
    If Dir(folder_path, vbDirectory) = "" Then
        MkDir folder_path
    End If
    
    ' Tex Data 書き込み
    Dim tex_code As String
    tex_code = "texinput_buf.tex"
    Dim file_path As String
    file_path = folder_path & "\" & tex_code
    Open file_path For Output As #1
    Print #1, UserForm1.TextBox1.Text
    Close #1

    ' =====================================================
    'texコンパイル
    Dim execCommand As String
    Dim wsh As Object
    Dim result As Integer
    '実行するコマンドを指定
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = folder_path
    
    'コマンドを同期実行
    execCommand = "ptex2pdf -u -l " & tex_code
    result = wsh.run(Command:="%ComSpec% /c " & execCommand, WaitOnReturn:=True)
    
    '後片付け
    Set wsh = Nothing
    ' =====================================================
    
    If (result = 0) Then
        Call pdf_to_svg
        ' ウィンドウを閉じる
        Unload UserForm1
    End If
End Sub

Private Sub UserForm_Initialize()
    With UserForm1
        .Caption = "TexInput"
    End With
    With TextBox1
        .MultiLine = True
        .ScrollBars = fmScrollBarsBoth
        .WordWrap = False
        .EnterKeyBehavior = True
        .Text = "\documentclass[uplatex,dvipdfmx]{jsarticle}" & vbNewLine & _
                "\usepackage{amsmath,amssymb,siunitx}" & vbNewLine & _
                "\usepackage{graphicx}" & vbNewLine & _
                "\usepackage{color}" & vbNewLine & _
                 vbNewLine & _
                "\newcommand{\ignore}[1]{}" & vbNewLine & _
                "\newcommand{\n}{\nonumber\\}" & vbNewLine & _
                "\pagestyle{empty}" & vbNewLine & _
                 vbNewLine & _
                "\begin{document}" & vbNewLine & _
                 vbNewLine & _
                "% ここに入力する" & vbNewLine & _
                 vbNewLine & _
                 vbNewLine & _
                "\end{document}"
    End With
    With CommandButton1
        .Caption = "実行"
    End With
End Sub

