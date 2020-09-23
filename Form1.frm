VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Do work with DoEvents"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4320
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   2400
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   240
      Width           =   495
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10%"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   20
         X1              =   16
         X2              =   16
         Y1              =   0
         Y2              =   100
      End
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin ProjectHandbrake.Duncan_Handbrake Duncan_Handbrake1 
      Left            =   1440
      Top             =   960
      _ExtentX        =   1270
      _ExtentY        =   1270
      CPULimit        =   30
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Do work with Handbrake"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "waiting?"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'results of testing
'64s on 20%
'62s on 50%


Private CPUrl(3) As Long

Private Sub Duncan_Handbrake1_CPUlevelCalculated(ByVal lVal As Long)
        CPUrl(2) = CPUrl(3)
        CPUrl(1) = CPUrl(2)
        CPUrl(0) = CPUrl(1)
        CPUrl(3) = lVal
End Sub

Private Sub Duncan_Handbrake1_ProcessingLoopTimed(ByVal lSeconds As Long)
    Label3 = lSeconds & " seconds processing"
End Sub

Private Sub Timer1_Timer()
    Dim lVal As Long
    lVal = (CPUrl(0) + CPUrl(1) + CPUrl(2) + CPUrl(3)) / 4
    DrawBar lVal
    Label1.Caption = lVal
    
    Label2.Caption = IIf(Duncan_Handbrake1.IsWaiting, "Waiting", "")
End Sub

Private Sub DrawBar(lSize As Long)
    Line1.Y1 = 100 - lSize
End Sub

Private Sub Command1_Click()
    Duncan_Handbrake1.Enabled = False
    Label2 = "canceled"
End Sub

Private Sub Command3_Click()
    Dim I As Long
    Dim D As Date
    D = Now
    For I = 1 To 10000
        DoEvents
        If StrComp("abcdefghijklmnopqrstuvwxyz0123456789", "zbcdefghijklmnopqrstuvwxyz0123456789", vbTextCompare) = 0 Then
        End If
        List1.AddItem "Itteration " & I, 0
    Next
    Label3 = DateDiff("s", D, Now) & " seconds processing"
End Sub

Private Sub Command2_Click()
    Dim I As Long
    List1.Clear
    Label1 = ""
    
    'turn timer on
    Duncan_Handbrake1.Enabled = True
    '0 inform entering a processing run
    Duncan_Handbrake1.InProcessingLoop = True
    For I = 1 To 10000
        '1a) check to see if unload event needs us to exit
        If Duncan_Handbrake1.UnloadInitiated Or (Not Duncan_Handbrake1.Enabled) Then
            Exit For
        End If
        DoEvents
        If StrComp("abcdefghijklmnopqrstuvwxyz0123456789", "zbcdefghijklmnopqrstuvwxyz0123456789", vbTextCompare) = 0 Then
        End If
        List1.AddItem "Itteration " & I, 0
        Duncan_Handbrake1.SlowProcessing
    Next
    'inform exiting from processing run
    'have this as the last command in the sub
    'dont have it nested way down. have it at top level
    'because when this command is called the app might be unloaded
    'so we dont want stuff after it
    Duncan_Handbrake1.InProcessingLoop = False
    Duncan_Handbrake1.Enabled = False
    If Duncan_Handbrake1.UnloadInitiated Then
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Duncan_Handbrake1.Enabled Then
        If Duncan_Handbrake1.InProcessingLoop Then
            Duncan_Handbrake1.UnloadInitiated = True
            Cancel = 1  'wait until handbrake is released
        End If
    End If

End Sub

