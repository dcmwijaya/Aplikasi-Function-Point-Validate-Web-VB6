VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000016&
   Caption         =   "CFP"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9675
   Icon            =   "CFP.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "CFP.frx":0A72
   ScaleHeight     =   7635
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      TabIndex        =   28
      Top             =   8280
      Width           =   3255
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H0000FF00&
      Caption         =   "Kembali ke menu utama"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8280
      Width           =   5655
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000FF00&
      Caption         =   "Kompleks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FF00&
      Caption         =   "Rata-rata"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000FF00&
      Caption         =   "Sederhana"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0000FF00&
      Caption         =   "Kompleks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0000FF00&
      Caption         =   "Rata-rata"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0000FF00&
      Caption         =   "Sederhana"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FF00&
      Caption         =   "Kompleks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "Rata-rata"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
      Caption         =   "Sederhana"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "Kompleks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "Rata-rata"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Sederhana"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Kompleks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Rata-rata"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0000FF00&
      Caption         =   "Nilai CFP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0000FF00&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Sederhana"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   4
      Top             =   6000
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   3
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   2
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   1
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "External Interface"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   26
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "File Logic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   25
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "User Inquery"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   24
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "User Output"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   23
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "User Input"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   22
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    IO.num1 = Val(Text1.Text) * Val(3)
End Sub
Private Sub Command2_Click()
    IO.num2 = Val(Text1.Text) * Val(4)
End Sub
Private Sub Command3_Click()
    IO.num3 = Val(Text1.Text) * Val(6)
End Sub
'################################################'
Private Sub Command4_Click()
    IO.num4 = Val(Text2.Text) * Val(4)
End Sub
Private Sub Command5_Click()
    IO.num5 = Val(Text2.Text) * Val(5)
End Sub
Private Sub Command6_Click()
    IO.num6 = Val(Text2.Text) * Val(7)
End Sub
'################################################'
Private Sub Command7_Click()
    IO.num7 = Val(Text3.Text) * Val(3)
End Sub
Private Sub Command8_Click()
    IO.num8 = Val(Text3.Text) * Val(4)
End Sub
Private Sub Command9_Click()
    IO.num9 = Val(Text3.Text) * Val(6)
End Sub
'################################################'
Private Sub Command10_Click()
    IO.num10 = Val(Text4.Text) * Val(7)
End Sub
Private Sub Command11_Click()
    IO.num11 = Val(Text4.Text) * Val(10)
End Sub
Private Sub Command12_Click()
    IO.num12 = Val(Text4.Text) * Val(15)
End Sub
'################################################'
Private Sub Command13_Click()
    IO.num13 = Val(Text5.Text) * Val(5)
End Sub
Private Sub Command14_Click()
    IO.num14 = Val(Text5.Text) * Val(7)
End Sub
Private Sub Command15_Click()
    IO.num15 = Val(Text5.Text) * Val(10)
End Sub
'################################################'
Private Sub Command16_Click()
    Dim NCFD As Double
    Text6.Text = IO.num1 + IO.num2 + IO.num3 + IO.num4 + IO.num5 + IO.num6 + IO.num7 + IO.num8 + IO.num9 + IO.num10 + IO.num11 + IO.num12 + IO.num13 + IO.num14 + IO.num15
    NCFD = Text6.Text
    IO.CFP = NCFD
End Sub
'################################################'
Private Sub Command17_Click()
    Form3.Show
End Sub
'################################################'
Private Sub Command18_Click()
    Form1.Show
End Sub
