VERSION 5.00
Begin VB.Form frmRemoveTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Tick Time"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   2760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox removetime 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Select time to remove:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmRemoveTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
main.Enabled = True
main.SetFocus
End Sub

Private Sub cmdremove_Click()
intLogedInAction = 4
frmRemoveTime.Enabled = False
frmLogin.Show
End Sub

Private Sub Form_Load()
Dim intCounter As Integer

intCounter = 0
For intCounter = 0 To CInt(intNumTicks)
    removetime.List(intCounter) = intTickTimes(intCounter)
Next
removetime.Text = intTickTimes(0)
End Sub
