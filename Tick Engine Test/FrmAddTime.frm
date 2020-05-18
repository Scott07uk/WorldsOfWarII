VERSION 5.00
Begin VB.Form FrmAddTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Time"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   2745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Ticktime 
      Height          =   315
      ItemData        =   "FrmAddTime.frx":0000
      Left            =   1800
      List            =   "FrmAddTime.frx":00B8
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Select a time to add:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAddTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
intLogedInAction = 3
FrmAddTime.Enabled = False
frmLogin.Show
End Sub

Private Sub cmdcancel_Click()
main.Enabled = True
Unload Me
main.SetFocus
End Sub

Private Sub Form_Load()
Ticktime.Text = 0
End Sub
