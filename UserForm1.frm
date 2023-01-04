VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6590
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7060
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    ' �t�H���_�쐬
    Dim folder_path As String
    folder_path = "C:\Windows\Temp\TexInput"
    If Dir(folder_path, vbDirectory) = "" Then
        MkDir folder_path
    End If
    
    ' Tex Data ��������
    Dim tex_code As String
    tex_code = "texinput_buf.tex"
    Dim file_path As String
    file_path = folder_path & "\" & tex_code
    Open file_path For Output As #1
    Print #1, UserForm1.TextBox1.Text
    Close #1

    ' =====================================================
    'tex�R���p�C��
    Dim execCommand As String
    Dim wsh As Object
    Dim result As Integer
    '���s����R�}���h���w��
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = folder_path
    
    '�R�}���h�𓯊����s
    execCommand = "ptex2pdf -u -l " & tex_code
    result = wsh.run(Command:="%ComSpec% /c " & execCommand, WaitOnReturn:=True)
    
    '��Еt��
    Set wsh = Nothing
    ' =====================================================
    
    If (result = 0) Then
        Call pdf_to_svg
        ' �E�B���h�E�����
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
                "% �����ɓ��͂���" & vbNewLine & _
                 vbNewLine & _
                 vbNewLine & _
                "\end{document}"
    End With
    With CommandButton1
        .Caption = "���s"
    End With
End Sub

