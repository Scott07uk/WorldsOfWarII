VERSION 5.00
Begin VB.Form frmProcessing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processing Tick"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2400
      Top             =   1680
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   4560
      Y1              =   1440
      Y2              =   960
   End
   Begin VB.Label lblbar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   1440
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The tick is now being processed, this may take some time."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCounter As Integer

Private Sub Form_Load()
intCounter = 1
End Sub

Private Sub Timer1_Timer()
'20
lblbar.Caption = lblbar.Caption & "@"

intCounter = intCounter + 1
If intCounter = 22 Then
    intCounter = 1
    lblbar.Caption = ""
End If

End Sub
