VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox passwordbox 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox usernamebox 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdlogin_Click()
Call processLogin
                        
End Sub

Private Sub Form_Load()
'clear any previous logins
usernamebox.Text = ""
passwordbox.Text = ""
End Sub


Public Sub processLogin()
If usernamebox.Text = "" Then
        MsgBox ("You have not entered a username.  ")
        usernamebox.SetFocus
    Else
        If passwordbox.Text = "" Then
                MsgBox ("You have not entered a password.  ")
                passwordbox.SetFocus
            Else
                
                main.SQLCon.RecordSource = "SELECT tblAccounts.admin FROM tblaccounts WHERE username = '" & Trim(usernamebox.Text) & "' AND password = '" & Trim(passwordbox.Text) & "';"
                
                main.SQLCon.Refresh
                
                If main.SQLCon.Recordset.EOF Then
                        MsgBox ("Login Failed")
                        passwordbox.Text = ""
                        usernamebox.Text = ""
                        usernamebox.SetFocus
                    Else
                        blnUserLogedin = True
                        Call returnProcessing
                        Unload Me
                End If
                
                'main.SQLCon.Recordset.Close
        End If
End If
End Sub

Private Sub passwordbox_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        Call processLogin
    Else
        If KeyCode = 27 Then
                passwordbox.Text = ""
        End If
End If
End Sub

Private Sub usernamebox_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        Call processLogin
    Else
        If KeyCode = 27 Then
                usernamebox.Text = ""
        End If
End If
End Sub
