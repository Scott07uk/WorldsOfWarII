VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Worlds of War II Tick Engine"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet EmailSender 
      Left            =   5520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc SQLCon2 
      Height          =   375
      Left            =   1920
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer TickTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   600
   End
   Begin VB.Frame frmgamestats 
      Caption         =   "Game Stats:"
      Height          =   2775
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   3375
      Begin VB.Label gamestats 
         Caption         =   "0"
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame currticktimes 
      Caption         =   "Tick Times:"
      Height          =   2415
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   1815
      Begin VB.Label lblticktimes 
         Caption         =   "0"
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdsettings 
      Caption         =   "Settings"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdremovetime 
      Caption         =   "Remove Time"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdaddTickTime 
      Caption         =   "Add Time"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdforcetick 
      Caption         =   "Force Tick"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "Stop Ticks"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start Ticks"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc SQLCon 
      Height          =   855
      Left            =   4680
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbltime 
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label status 
      Alignment       =   2  'Center
      Caption         =   "Worlds of War II Tick Service Stopped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaddTickTime_Click()
FrmAddTime.Show
main.Enabled = False
End Sub

Private Sub cmdforcetick_Click()
intLogedInAction = 2
frmLogin.Show
main.Enabled = False
End Sub

Private Sub cmdremovetime_Click()
main.Enabled = False
frmRemoveTime.Show
End Sub

Private Sub cmdsettings_Click()
main.Enabled = True
frmSettings.Show
End Sub

Private Sub cmdstart_Click()
'make sure we can start the ticks
If blnCanTicksStart = True Then
        'sync the tick counter
        Call syncTimer
        TickTimer.Enabled = True
        cmdstart.Enabled = False
        cmdstop.Enabled = True
        status.Caption = "Ticks Running"
        status.ForeColor = RGB(0, 255, 0)
    Else
        MsgBox ("Ticks can not start as there are no times defined")
End If
End Sub

Private Sub cmdstop_Click()
TickTimer.Enabled = False
cmdstart.Enabled = True
cmdstop.Enabled = False
End Sub

Private Sub Form_Load()
'check to see if a config file exists
strSettingsFile = App.Path & "\settings.txt"
'open the settinfs file
Open strSettingsFile For Input As 1
'get the settings from the settings file
Input #1, blnAuditOn
'close the settings file
Close #1

'clear anything in the game status
gamestats.Caption = ""

'set the connection string
SQLCon.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWTick;Password=therealtickengine;Database=WoWII;"

SQLCon2.ConnectionString = "Provider=SQLOLEDB;Server=neptune;User ID=WoWTick;Password=therealtickengine;Database=WoWII;"

SQLCon.RecordSource = "SELECT * FROM TblTickTimes;"

SQLCon.Refresh

If SQLCon.Recordset.EOF Then
        blnCanTicksStart = False
        lblticktimes.Caption = "NONE SET"
    Else
        blnCanTicksStart = True
        
        'load the tick times
        intLoopCounter = 0
        lblticktimes.Caption = ""
        Do While Not SQLCon.Recordset.EOF
            intTickTimes(intLoopCounter) = SQLCon.Recordset![Ticktime]
            intLoopCounter = intLoopCounter + 1
            lblticktimes.Caption = lblticktimes.Caption & SQLCon.Recordset![Ticktime] & vbCrLf
            
            SQLCon.Recordset.MoveNext
        Loop
        
        'set the number of ticks to the number of times the loop was ran
        intNumTicks = intLoopCounter - 1
End If

'close the recordset
SQLCon.Recordset.Close

'loading is done

End Sub

Private Sub TickTimer_Timer()
Dim intCounter As Integer

intCounter = 0
intTimerSeconds = intTimerSeconds + 1

If intTimerSeconds = 60 Then
        intTimerSeconds = 0
        intTimerMinutes = intTimerMinutes + 1
End If

If intTimerMinutes = 60 Then
        intTimerHours = intTimerHours + 1
        intTimerMinutes = 0
End If

If intTimerHours = 24 Then
        intTimerHours = 0
        Call syncTimer
End If

lbltime.Caption = "Time:" & intTimerHours & ":" & intTimerMinutes & ":" & intTimerSeconds

If (intTimerSeconds = 30) Then
        Call EmailSender.OpenURL("http://www.worldsofwar.co.uk/purge_email.asp")
End If

For intCounter = 0 To intNumTicks
    If intTimerMinutes = intTickTimes(intCounter) Then
            'make sure the tick is only called once
            If intTimerSeconds = 0 Then
                    Call processTick
                    Call syncTimer
            End If
    End If
Next

End Sub
