VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CLICK ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CLICK ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Made by RYAN HOLLETT FEB 23 2000
' HIGHLIGHTING LABELS

Private Sub Command1_Click()
    MsgBox "You have changed my color"
End Sub
Private Sub label1_Click()
    MsgBox "You have changed my color"
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Command1.BackColor = vbRed
    Let Label1.BackColor = vbYellow
End Sub
Private Sub Form_Load()
    Let Command1.BackColor = vbYellow
    Let Label1.BackColor = vbYellow
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Command1.BackColor = vbYellow
    Let Label1.BackColor = vbYellow
End Sub
Private Sub label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let Label1.BackColor = vbRed
    Let Command1.BackColor = vbYellow
End Sub

