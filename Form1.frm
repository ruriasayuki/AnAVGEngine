VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "test"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "����"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5400
   ScaleMode       =   0  'User
   ScaleWidth      =   8100
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "������κ"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   3495
      Width           =   75
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   550
      TabIndex        =   0
      Top             =   3960
      Width           =   7000
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   0
      Picture         =   "Form1.frx":437C
      Top             =   0
      Visible         =   0   'False
      Width           =   8100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API�����Ķ���
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength _
As Long, ByVal hwndCallback As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1&
Dim s, t As String
Dim k, lk, lk_1 As Boolean '����main���� �Լ��������
Dim p, q As String
Dim X As String  '����changeָ��
'���̲����Ķ��� �س��������� �ո����ضԻ��� ��ctrl
Private Sub Form_keydown(keycode As Integer, Shift As Integer)
If keycode = 13 And lk Then main
If keycode = 32 Then hidebox
If keycode = 17 And lk Then Timer1.Enabled = True
End Sub
Private Sub Form_keyup(keycode As Integer, Shift As Integer)
If keycode = 32 Then showbox
If keycode = 17 Then Timer1.Enabled = False
End Sub
'�򿪽ű�
Private Sub op(a As String)
Dim b As String
    Close #1
    b = "txt\" + a + ".avg"
    Open b For Input As #1
    main
End Sub
'��ʼ��
Public Sub Form_Load()
    Timer1.Enabled = False
    Image1.Visible = False
    Form1.Picture = LoadPicture("")
    X = "1"
    op (X)
    lk = True
    lk_1 = True
End Sub
'����ִ�нű�
Public Sub main()
Dim i As Integer
Line Input #1, s
    t = ""
    p = Mid(s, 1, 1)
    If p = "\" Then
        k = False
        For i = 1 To Len(s) Step 1
            q = Mid(s, i, 1)
            If q = " " Then
                t = Mid(s, 2, i - 2)
                s = Mid(s, i + 1, Len(s) - i)
                k = True
                Exit For
            End If
        Next i
        If Not k Then
            t = Mid(s, 2, Len(s) - 1)
            s = ""
        End If
        '�����ű���ͷָ��
        Select Case t
        Case "bg"
            bg (s) '����ͼƬ ���� ͼƬ����
            main
        Case "bgm"
            bgm (s) '�������� ���� ��������
            main
        Case "ch" '���� ���� ��������
            ch (s)
            main
        Case "vo" '������������ ���� ��������
            vo (s)
            main
        Case "se" '��������Ч�� ���� ��Ч����
            se (s)
            main
        Case "op" '���½ű� ���� �ű�����
            op (s)
        Case "cl" '�������� �޲���
            cl
            main
        Case "clall" '���� �޲���
            clall
            main
        Case "end" '�������˳��� �޲���
            ed
        Case "sel" 'ѡ��֫ ���� ѡ��֫������1~4��
            sel (s)
        Case "goto" 'goto��� ���� ��ǩ��
            gt (s)
        Case "wait" 'ʱ�� ���� ʱ�䳤�� ���루str��
            wait (s)
        Case "auto" '�Զ�ִ�� ���� ʱ�䳤��
            auto (s)
        End Select
    Else
    txt (s)
    End If
End Sub
'����ͼƬ����ָ��
Private Sub bg(a As String)
Dim b As String
    b = "bg\" + a + ".jpg"
    Form1.Picture = LoadPicture(b)
End Sub
'�������ָ���ָ��
Private Sub bgm(a As String)
Dim b As String
    mciSendString "close all", vbNullString, 0, 0
    If a <> "" Then
    b = "bgm\" + a + ".mp3"
    Call mciSendString("play " & b & " repeat", vbNullString, 0, 0)
    End If
End Sub
'���棨��ʱֻ֧�ֵ��ţ�ָ��
Private Sub ch(a As String)
Image1.Visible = False
Dim b, c, d, e, f As String
Dim i, h, z As Integer
b = a
For i = 1 To Len(a)
    If Mid(a, i, 1) = " " Then
        b = Mid(a, 1, i - 1)
        c = Mid(a, i + 1, Len(a) - i)
        Exit For
    End If
Next i
e = c
For i = 1 To Len(c)
    If Mid(c, i, 1) = " " Then
        e = Mid(c, 1, i - 1)
        f = Mid(c, i + 1, Len(c) - i)
        Exit For
    End If
Next i
    d = "chara\" + b + ".gif" '��Ҫ�ڴ���߼�����
    h = Val(e)
    z = Val(f)
    Image1.Visible = False
    Image1.Picture = LoadPicture(d)
    Image1.Left = Int(Form1.Width * h / 100)
    Image1.Top = Int(Form1.Height * z / 100)
    Image1.Visible = True
End Sub
'�Ի�����ָ��
Private Sub txt(a As String)
Dim t, f, n, b As String
Dim i As Integer
Label1.Visible = True
    b = a
    t = Mid(a, 1, 1)
    If t = "{" Then
    For i = 2 To Len(a)
        f = Mid(a, i, 1)
        If f = "}" Then
           n = Mid(a, 2, i - 2)
           b = Mid(a, i + 1, Len(a) - i)
           na (n)
        End If
        Next i
    Else
        na (n)
    End If
    Label1.Caption = b
End Sub
'����ָ��
Private Sub vo(a As String)
Dim b As String
    b = "voice\" + a + ".wav" 'ͬ�� Ӧ�������Զ���·��
    Call sndPlaySound(b, SND_ASYNC)
End Sub
'Ч����ָ�ͬ��������ʱ�Ҳ���ͬʱ�����췽��
Private Sub se(a As String)
    b = "se\" + a + ".wav" 'ͬ�� �Զ���·������չ
    Call sndPlaySound(b, SND_ASYNC)
End Sub
'��������ж�
Private Sub na(a As String)
If a = "" Then
    Label2.Visible = False
Else
    Label2.Visible = True
    Label2.Caption = a
    Label2.BackStyle = 1
    Label2.BorderStyle = 1
End If
End Sub
'�������ָ��
Private Sub cl()
    Image1.Picture = LoadPicture("")
End Sub
'����ָ��
Private Sub clall()
    cl
    Label1.Visible = False
    Label2.Visible = False
End Sub
'ѡ��ָ֫��(������)
Private Sub sel(a)
lk = False
Dim i As Integer
    Timer1.Enabled = False
    Timer2.Enabled = False
    na ("")
    st = Val(a)
    For i = 1 To 4
        wenjian(i) = ""
    Next i
    For i = 1 To st
        Line Input #1, wenjian(i)
    Next i
    Label1.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
    Label3.Caption = wenjian(1)
    Label4.Caption = wenjian(2)
    Label5.Caption = wenjian(3)
    Label6.Caption = wenjian(4)
    If Label3.Caption <> "" Then Label3.Visible = True
    If Label4.Caption <> "" Then Label4.Visible = True
    If Label5.Caption <> "" Then Label5.Visible = True
    If Label6.Caption <> "" Then Label6.Visible = True
End Sub
'ѡ��ָ֫�������֣�
Private Sub Label3_Click()
temp = 1
selend
main
End Sub

Private Sub Label4_Click()
temp = 2
selend
main
End Sub
Private Sub Label5_Click()
temp = 3
selend
main
End Sub
Private Sub Label6_Click()
temp = 4
selend
main
End Sub
'ѡ��ָ֫�����ѡ��
Private Sub selend()
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label3.BackStyle = 0
    Label4.BackStyle = 0
    Label5.BackStyle = 0
    Label6.BackStyle = 0
    lk = True
End Sub
'ѡ��ָ֫������ͣЧ����
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.Visible Then
    Label3.BackStyle = 1
    Label3.BackColor = &HFFFF00
    Label4.BackStyle = 0
    Label5.BackStyle = 0
    Label6.BackStyle = 0
End If
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label4.Visible Then
    Label4.BackStyle = 1
    Label4.BackColor = &HFFFF00
    Label3.BackStyle = 0
    Label5.BackStyle = 0
    Label6.BackStyle = 0
End If
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.Visible Then
    Label5.BackStyle = 1
    Label5.BackColor = &HFFFF00
    Label3.BackStyle = 0
    Label4.BackStyle = 0
    Label6.BackStyle = 0
End If
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.Visible Then
    Label6.BackStyle = 1
    Label6.BackColor = &HFFFF00
    Label3.BackStyle = 0
    Label4.BackStyle = 0
    Label5.BackStyle = 0
End If
End Sub
'���ضԻ��� ���¿ո������
Private Sub hidebox()
    Label1.Visible = False
    Label2.Visible = False
    selend
End Sub
'��ʾ�Ի��� �ɿ��ո������
Private Sub showbox()
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
End Sub
'gotoָ��
Private Sub gt(a As String)
Dim i, m As Integer
Dim b, c As String
t = ""
b = a
c = ""
For i = 1 To Len(a)
    t = Mid(a, i, 1)
    If t = " " Then
        b = Mid(a, 1, i - 1)
        c = Mid(a, i + 1, Len(a) - i)
        Exit For
    End If
Next i
m = Val(c)
b = "\" + b
If m = temp Or c = "" Then
    Line Input #1, s
    While s <> b
    Line Input #1, s
    Wend
End If
main
End Sub
'waitָ�� ��Timer2���� �Ǳ����������õ���ʱ������Ҫ�Ŀؼ�
Private Sub wait(a As String)
Dim b As Integer
If Not (Timer1.Enabled) Then
    lk = False
    b = Val(a)
    Timer2.Interval = b
    Timer2.Enabled = True
End If
End Sub
'waitָ����ӳ��� ����ʵ���ӳ�
Private Sub Timer2_Timer()
    Timer2.Enabled = False
    lk = True
    main
End Sub
'�Զ�ִ����һ����ָ���Ϊ�Ի�����ָ��
Private Sub auto(a As String)
    main
    wait (a)
End Sub
'�ļ�β��֤
Private Sub ed()
    Close #1
    lk = False
    bgm ("")
    Form1.Hide
    Form2.Show
End Sub
'���ָ�� ��ctrlָ���
Private Sub Timer1_Timer()
    If lk Then main
End Sub
