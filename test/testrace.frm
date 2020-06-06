VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "中山競馬場レース場 Ver7"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows の既定値
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   7440
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   2535
   End
   Begin VB.OptionButton Option5 
      Caption         =   "フランソワインター"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   6480
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      Caption         =   "ガマゴオリサンシー"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   6000
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "ヨジマンジャー"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   5520
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "シカマウリスター"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "エイコサペンター"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Let's Start!!!"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   4320
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   12915
      TabIndex        =   0
      Top             =   0
      Width           =   12975
   End
   Begin VB.Label Label6 
      Height          =   615
      Left            =   1560
      TabIndex        =   15
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "所持金"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   12
      Top             =   5520
      Width           =   6255
   End
   Begin VB.Label Label3 
      Caption         =   "Hikaru"
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   6
      Top             =   4440
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "ゴールだ↑"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, n, m, a, b, c, d, h, i, j, k, l, o, p, q, x1, y1, n1, m1, z, u, aa､bb, o1, o2, o3, o4, o5 As Single

Private Sub Command1_Click()
z = 1
List1.Clear
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False Then
    u = MsgBox("馬を予想してくれ!!!", vbOKOnly, "エラー") = vbOK
    z = 0
End If
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Option5.Visible = False
If Option1.Value = True Then
    Label4.Caption = "あなたの予想は" & Option1.Caption & "（黄）"
ElseIf Option2.Value = True Then
    Label4.Caption = "あなたの予想は" & Option2.Caption & "（赤）"
ElseIf Option3.Value = True Then
    Label4.Caption = "あなたの予想は" & Option3.Caption & "（緑)"
ElseIf Option4.Value = True Then
    Label4.Caption = "あなたの予想は" & Option4.Caption & "（青）"
ElseIf Option5.Value = True Then
    Label4.Caption = "あなたの予想は" & Option5.Caption & "（紫）"
End If
End Sub


Private Sub Command2_Click()
Picture1.Cls
o1 = 0
o2 = 0
o3 = 0
o4 = 0
o5 = 0
z = 0
X = 100
Y = 300
n = 0
m = 0
a = 100
b = 1200
c = 0
d = 0
h = 100
i = 2100
j = 0
k = 0
l = 100
p = 3000
o = 0
q = 0
x1 = 100
y1 = 3900
n1 = 0
m1 = 0
Label2.Caption = ""
Label4.Caption = ""
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True
Option4.Visible = True
Option5.Visible = True
End Sub

Private Sub Form_Load()
Timer1.Interval = 10
X = 100
Y = 300
n = 0
m = 0
a = 100
b = 1200
c = 0
d = 0
h = 100
i = 2100
j = 0
k = 0
l = 100
p = 3000
o = 0
q = 0
x1 = 100
y1 = 3900
n1 = 0
m1 = 0
aa = Val(Text1.Text)
bb = 1000
o1 = 0
o2 = 0
o3 = 0
o4 = 0
o5 = 0
Label6.Caption = bb & "円"
End Sub




Private Sub Timer1_Timer()
If z = 1 Then
    Randomize
    Picture1.Line (0, 0)-(14000, 9000), , BF
    aa = Val(Text1.Text)
    bb = 1000
    Text1.Text = Val(bb - aa)
    Label6.Caption = bb - aa & "円"
    X = X + n
    Y = Y + m
    a = a + c
    b = b + d
    h = h + j
    i = i + k
    l = l + o
    p = p + q
    x1 = x1 + n1
    y1 = y1 + m1
    n = Rnd * 80
    c = Rnd * 80
    j = Rnd * 80
    o = Rnd * 80
    n1 = Rnd * 80
    Picture1.AutoRedraw = True
    Picture1.FillStyle = 0
    Picture1.Line (12000, 0)-(12000, Picture1.Height), RGB(255, 0, 0)
    Picture1.Line (400, 0)-(400, Picture1.Height), RGB(0, 0, 255)
    Picture1.FillColor = RGB(255, 255, 0)
    Picture1.Circle (X, Y), 100, RGB(255, 255, 0)
    Picture1.FillColor = RGB(255, 0, 0)
    Picture1.Circle (a, b), 100, RGB(255, 0, 0)
    Picture1.FillColor = RGB(0, 255, 0)
    Picture1.Circle (h, i), 100, RGB(0, 255, 0)
    Picture1.FillColor = RGB(0, 0, 255)
    Picture1.Circle (l, p), 100, RGB(0, 0, 255)
    Picture1.FillColor = RGB(255, 0, 255)
    Picture1.Circle (x1, y1), 100, RGB(255, 0, 255)
    If X >= 12000 Then
        If Option1.Value = True Then
            Label2.Caption = "おめでとう!!!!!"
        Else
            Label2.Caption = "残念!!!"
        End If
        If o1 = 0 Then
            List1.AddItem Option1.Caption & "（黄）"
            o1 = 1
        End If
    End If
    If a >= 12000 Then
        If Option2.Value = True Then
            Label2.Caption = "おめでとう!!!!!"
        Else
            Label2.Caption = "残念!!!"
        End If
        If o2 = 0 Then
            List1.AddItem Option2.Caption & "（赤）"
            o2 = 1
        End If
    End If
    If h >= 12000 Then
        If Option3.Value = True Then
            Label2.Caption = "おめでとう!!!!!"
        Else
            Label2.Caption = "残念!!!"
        End If
        If o3 = 0 Then
            List1.AddItem Option3.Caption & "（緑）"
            o3 = 1
        End If
    End If
    If l >= 12000 Then
        If Option4.Value = True Then
            Label2.Caption = "おめでとう!!!!!"
        Else
            Label2.Caption = "残念!!!"
        End If
        If o4 = 0 Then
            List1.AddItem Option4.Caption & "（青）"
            o4 = 1
        End If
    End If
    If x1 >= 12000 Then
        If Option5.Value = True Then
            Label2.Caption = "おめでとう!!!!!"
        Else
            Label2.Caption = "残念!!!"
        End If
        If o5 = 0 Then
            List1.AddItem Option5.Caption & "（紫）"
            o5 = 1
        End If
    End If
        If X < (a And h And l And x1) Then
        Label2.Caption = "エイコサペンターが速いぞ"
    ElseIf a < (X And h And l And x1) Then
        Label2.Caption = "シカマウリスターが速いぞ"
    ElseIf h < (X And a And l And x1) Then
        Label2.Caption = "ヨジマンジャーが速いぞ"
    ElseIf l < (X And a And h And x1) Then
        Label2.Caption = "ガマゴオリサンシーが速いぞ"
    ElseIf x1 < (X And a And h And l) Then
        Label2.Caption = "フランソワインターが速いぞ"
    End If
End If
End Sub
