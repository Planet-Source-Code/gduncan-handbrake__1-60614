VERSION 5.00
Begin VB.UserControl Duncan_Handbrake 
   BackColor       =   &H00FF80FF&
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Duncan_Handbrake.ctx":0000
   ScaleHeight     =   1560
   ScaleWidth      =   3525
   ToolboxBitmap   =   "Duncan_Handbrake.ctx":0ECA
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1440
      Top             =   240
   End
End
Attribute VB_Name = "Duncan_Handbrake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'What is this?
'A timer that is used to slow processing.

'Why?
'I had an app that needed to do database maintenance to a very large number
'of records.
'Firing the maintence sub caused the cpu to run at 100% for several minutes
'during which time the user could not do anything as their system all
'but locked up.
'Hence this handbrake. It does a burst, then waits, then a burst, then waits
'and the user is uneffected by lag.

'How?
'If you set CPULimit to 0 it will force a 1 second wait through each pass
'of your loop which in most cases should be enough to prevent the system
'from getting locked up under heavy number crunching.
'Alternativly you can use CPULimit to set the top end at which CPU useage
'can be before "waits" are introduced.
'eg process as hard as you like but as soon as 50% processor power is
'used then stop/wait because you are demanding too much of the system.


'Warning!
'Must be used properly!
'Because the form continues to receive and process messages while we are "waiting"
'it is possible the application may try to close while we are in a "wait" state
'This woud be bad because the app would process the Form_Unload event
'and then our timer would kick off and we would be back in the sub we called it from
'and our app would still be running
'So...to avoid this it is important that the loop this control is
'embedded within is structured properly as is the form_cancel sub.
'see sample app for details

'Who?
'Wait Timer code sourced from http://support.microsoft.com/default.aspx?scid=http://support.microsoft.com:80/support/kb/articles/Q231/2/98.ASP&NoWebContent=1
'CPU useage code sourced from http://www.allapi.net

'When?
'Last Updated : May 2005

'Bugs
'The cancel timer doesnt work - cant understand why. Pls help if you can.

'Note: I only put the code in for CPU monitoring
'for systems of NT, 2000, and XP
'if you want CPU monitoring for 98 or below then you can add it
'in but I had no need for it. Source can be found at http://www.allapi.net

'Note 2 : Sampling speed of the timer is very important.
'150 seems to work about right for me.

'======================================================================================================================================================
'MY DECLARES FOR THIS CONTROL
'======================================================================================================================================================
'------------
'API DECLARES
'------------
'for Wait Timer
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&

Private Const INFINITE = &HFFFF
Private Const ERROR_ALREADY_EXISTS = 183&

Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT _
                            Or QS_POSTMESSAGE _
                            Or QS_TIMER _
                            Or QS_PAINT _
                            Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE _
                            Or QS_PAINT _
                            Or QS_TIMER _
                            Or QS_POSTMESSAGE _
                            Or QS_MOUSEBUTTON _
                            Or QS_MOUSEMOVE _
                            Or QS_HOTKEY _
                            Or QS_KEY)

Private Declare Function CreateWaitableTimer Lib "kernel32" _
    Alias "CreateWaitableTimerA" ( _
    ByVal lpSemaphoreAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function OpenWaitableTimer Lib "kernel32" _
    Alias "OpenWaitableTimerA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function SetWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long, _
    lpDueTime As FILETIME, _
    ByVal lPeriod As Long, _
    ByVal pfnCompletionRoutine As Long, _
    ByVal lpArgToCompletionRoutine As Long, _
    ByVal fResume As Long) As Long
    
Private Declare Function CancelWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long)
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Function MsgWaitForMultipleObjects Lib "user32" ( _
    ByVal nCount As Long, _
    pHandles As Long, _
    ByVal fWaitAll As Long, _
    ByVal dwMilliseconds As Long, _
    ByVal dwWakeMask As Long) As Long

'for NT CPU monitor
Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0
Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type
Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type
Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias  As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private liOldIdleTime As LARGE_INTEGER
Private liOldSystemTime As LARGE_INTEGER

'-----------
'MY DECLARES
'-----------
Private m_hTimer As Long                'the timer
Private m_IsWaiting As Boolean          'how we know what state we are in
Private m_InProcessingLoop As Boolean   'how we restrict the app from closing when a wait is happening
Private m_Enabled As Boolean            'how we cancel out
Private m_UnloadInitiated As Boolean    'how we know if unload is needed
Private m_CPU As Long                   'average useage of the CPU in the last second
Private m_CPULimit As Long              'cpu cycles must be below this or wait is initiated
Private m_SecondsToWait As Long         'how long each "wait" should be
Private m_StartTime As Date             'so we can calc time taken
Public Event CPUlevelCalculated(ByVal lVal As Long)
Public Event ProcessingLoopTimed(ByVal lSeconds As Long)

'======================================================================================================================================================
'                               CODE STARTS
'======================================================================================================================================================

'-----------------
'PUBLIC PROPERTIES
'-----------------
Public Property Get IsWaiting() As Boolean
    'is a "wait" in process?
    IsWaiting = m_IsWaiting
End Property

Public Property Get InProcessingLoop() As Boolean
Attribute InProcessingLoop.VB_MemberFlags = "400"
    'set this at the start and end of a processing loop so we know you have
    'exited ok
    InProcessingLoop = m_InProcessingLoop
End Property
Public Property Let InProcessingLoop(bVal As Boolean)
    If bVal Then
        m_StartTime = Now
    Else
        RaiseEvent ProcessingLoopTimed(DateDiff("s", m_StartTime, Now))
    End If
    m_InProcessingLoop = bVal
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(bVal As Boolean)
    m_Enabled = bVal
    Timer1.Enabled = m_Enabled
End Property

Public Property Get UnloadInitiated() As Boolean
Attribute UnloadInitiated.VB_MemberFlags = "400"
    UnloadInitiated = m_UnloadInitiated
End Property
Public Property Let UnloadInitiated(bVal As Boolean)
    m_UnloadInitiated = bVal
End Property

Public Property Get CPULimit() As Long
    CPULimit = m_CPULimit
End Property
Public Property Let CPULimit(lVal As Long)
    m_CPULimit = lVal
End Property

Public Property Get SecondsToWait() As Long
    If m_SecondsToWait <= 0 Then m_SecondsToWait = 1
    SecondsToWait = m_SecondsToWait
End Property
Public Property Let SecondsToWait(lVal As Long)
    m_SecondsToWait = lVal
End Property


'----------------
'PUBLIC FUNCTIONS
'----------------
Public Sub SlowProcessing()
    'are we going too fast?
    If Enabled Then
        If CPULimit > 0 Then
            'we test using processor speed
            If m_CPU > CPULimit Then
                Debug.Print "pausing " & m_CPU & ">" & CPULimit
            Else
                Debug.Print "skipping " & m_CPU & "<" & CPULimit
            End If
            Do While (m_CPU > CPULimit) And Enabled And (Not UnloadInitiated)
                m_IsWaiting = True
                Wait SecondsToWait      'wait for a bit
                m_IsWaiting = False
            Loop
            'Debug.Print m_CPU & "<" & CPULimit; " so not waiting"
        Else
            'we force wait
            m_IsWaiting = True
            Wait 1      'wait for a bit
            m_IsWaiting = False
        End If
    End If
End Sub

Public Sub CancelWait()
    'should break you out of the "wait" so that you dont have to wait
    'for its whole duration but it doesnt seem to work !!!!!
    'Always fails. Useless p.o.s
    '
    On Error Resume Next
    Dim retval As Long
    
    If IsWaiting Then
        If m_hTimer <> 0 Then
            'open the handle to the timer
            'retval = OpenWaitableTimer(TIMER_MODIFY_STATE, 0, m_TimerName)
            'If retval <> 0 Then
            'opened ok
            'tell it to cancel
            retval = CancelWaitableTimer(m_hTimer)
            If retval = 0 Then
                Debug.Print "canceling failed"
                'failed
            End If
            'End If
        
            'close it
            retval = CloseHandle(m_hTimer)
            If retval = 0 Then
                'failed
                Debug.Print "Closing handle failed"
            Else
                m_hTimer = 0
            End If
        End If
    End If
End Sub

'-----------------
'PRIVATE FUNCTIONS
'-----------------
Private Sub Wait(lNumberOfSeconds As Long)
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    Dim dblUnits As Double
    
    m_hTimer = CreateWaitableTimer(0, True, App.EXEName & Extender.Name)
    
    If Err.LastDllError = ERROR_ALREADY_EXISTS Then
        ' If the timer already exists, it does not hurt to open it
        ' as long as the person who is trying to open it has the
        ' proper access rights.
    Else
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(m_hTimer, ft, 0, 0, 0, 0)
    End If
    
    ' Convert the Units to nanoseconds.
    dblUnits = CDbl(&H10000) * CDbl(&H10000)
    dblDelay = CDbl(lNumberOfSeconds) * 1000 * 10000
    
    ' By setting the high/low time to a negative number, it tells
    ' the Wait (in SetWaitableTimer) to use an offset time as
    ' opposed to a hardcoded time. If it were positive, it would
    ' try to convert the value to GMT.
    ft.dwHighDateTime = -CLng(dblDelay / dblUnits) - 1
    dblDelayLow = -dblUnits * (dblDelay / dblUnits - _
        Fix(dblDelay / dblUnits))
    
    If dblDelayLow < CDbl(&H80000000) Then
        ' &H80000000 is MAX_LONG, so you are just making sure
        ' that you don't overflow when you try to stick it into
        ' the FILETIME structure.
        dblDelayLow = dblUnits + dblDelayLow
    End If
    
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(m_hTimer, ft, 0, 0, 0, False)
    
    Do
        If m_hTimer = 0 Then Exit Do
        If m_Enabled = False Then Exit Do
        ' QS_ALLINPUT means that MsgWaitForMultipleObjects will
        ' return every time the thread in which it is running gets
        ' a message. If you wanted to handle messages in here you could,
        ' but by calling Doevents you are letting DefWindowProc
        ' do its normal windows message handling---Like DDE, etc.
        lBusy = MsgWaitForMultipleObjects(1, m_hTimer, False, _
            INFINITE, QS_ALLINPUT&)
        DoEvents
        'Debug.Print Now
    Loop Until lBusy = WAIT_OBJECT_0
    
    'Clean up
    If m_hTimer <> 0 Then
        CloseHandle m_hTimer
        m_hTimer = 0
    End If

End Sub

'------------
'USER CONTROL
'------------
Private Sub UserControl_InitProperties()
    SecondsToWait = 1
    CPULimit = 10
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Debug.Print "control loaded " & Now
    CPULimit = PropBag.ReadProperty("CPULimit", 0)
    SecondsToWait = PropBag.ReadProperty("SecondsToWait", 1)
    InitialiseCPU
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CPULimit", CPULimit, 0)
    Call PropBag.WriteProperty("SecondsToWait", SecondsToWait, 1)
    
End Sub

Private Sub UserControl_Resize()
    If Not Ambient.UserMode Then
        UserControl.BackColor = UserControl.Parent.BackColor
        UserControl.Width = 48 * Screen.TwipsPerPixelX
        UserControl.Height = 48 * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub UserControl_Terminate()
    Dim retval As Long
    Timer1.Enabled = False
    If m_hTimer <> 0 Then
        CloseHandle m_hTimer
        m_hTimer = 0
    End If
    Debug.Print "unloaded"
End Sub


'-----
'TIMER
'-----
Private Sub Timer1_Timer()
    'sets m_CPU to the average of the last 4 sampled cpu useage values
    'this is compared against CPULimit to see if we need to slow processing
    'take alot of samples so that we get a good reading
    Dim Ret As Long
    Static CPUrl(9) As Long 'CPU recorded level
    
    'query the CPU usage
    Ret = QueryCPU
    If Ret = -1 Then
        Debug.Print "CPU testing is unavailable"
        Enabled = False
    Else
        CPUrl(8) = CPUrl(9)
        CPUrl(7) = CPUrl(8)
        CPUrl(6) = CPUrl(7)
        CPUrl(5) = CPUrl(6)
        CPUrl(4) = CPUrl(5)
        CPUrl(3) = CPUrl(4)
        CPUrl(2) = CPUrl(3)
        CPUrl(1) = CPUrl(2)
        CPUrl(0) = CPUrl(1)
        CPUrl(9) = Ret
        m_CPU = (CPUrl(0) + CPUrl(1) + CPUrl(2) + CPUrl(3) + CPUrl(4) + CPUrl(5) + CPUrl(6) + CPUrl(7) + CPUrl(8) + CPUrl(9)) / 10
        RaiseEvent CPUlevelCalculated(m_CPU)
    End If
End Sub

'-------------------
'CPU MONITORING SUBS
'-------------------
'This code from The KPD-Team at http://www.allapi.net
Private Sub InitialiseCPU()
    Dim SysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim SysPerfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim Ret As Long
    'get new system time
    Ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while initializing the system's time!", vbCritical
        Exit Sub
    End If
    'get new CPU's idle time
    Ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), ByVal 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while initializing the CPU's idle time!", vbCritical
        Exit Sub
    End If
    'store new CPU's idle and system time
    liOldIdleTime = SysPerfInfo.liIdleTime
    liOldSystemTime = SysTimeInfo.liKeSystemTime
End Sub
Public Function QueryCPU() As Long
    Dim SysBaseInfo As SYSTEM_BASIC_INFORMATION
    Dim SysPerfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim SysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim dbIdleTime As Currency
    Dim dbSystemTime As Currency
    Dim Ret As Long
    QueryCPU = -1
    'get number of processors in the system
    Ret = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(SysBaseInfo), LenB(SysBaseInfo), 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while retrieving the number of processors!", vbCritical
        Exit Function
    End If
    'get new system time
    Ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while retrieving the system's time!", vbCritical
        Exit Function
    End If
    'get new CPU's idle time
    Ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), ByVal 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while retrieving the CPU's idle time!", vbCritical
        Exit Function
    End If
    'CurrentValue = NewValue - OldValue
    dbIdleTime = LI2Currency(SysPerfInfo.liIdleTime) - LI2Currency(liOldIdleTime)
    dbSystemTime = LI2Currency(SysTimeInfo.liKeSystemTime) - LI2Currency(liOldSystemTime)
    'CurrentCpuIdle = IdleTime / SystemTime
    If dbSystemTime <> 0 Then dbIdleTime = dbIdleTime / dbSystemTime
    'CurrentCpuUsage% = 100 - (CurrentCpuIdle * 100) / NumberOfProcessors
    dbIdleTime = 100 - dbIdleTime * 100 / SysBaseInfo.bKeNumberProcessors + 0.5
    QueryCPU = Int(dbIdleTime)
    'store new CPU's idle and system time
    liOldIdleTime = SysPerfInfo.liIdleTime
    liOldSystemTime = SysTimeInfo.liKeSystemTime
End Function
Private Function LI2Currency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LI2Currency, liInput, LenB(liInput)
End Function

