VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm_Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " GMLAN Testing Program - v0.1"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "frm_Main"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   7800
      TabIndex        =   49
      Top             =   2520
      Width           =   2055
      Begin VB.TextBox txt_Buffer 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmd_ShowLog 
         Caption         =   "Show Incoming Data"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbl_Shortcuts 
         Alignment       =   2  'Center
         Caption         =   "F10 - Toggle KB Shortcuts"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "me@andrewgreen.ca"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "March 2013 - Andrew G"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2280
         Width           =   1815
      End
   End
   Begin VB.Frame frm_KeyFob 
      Caption         =   "Key Fob Buttons"
      Height          =   2895
      Left            =   5160
      TabIndex        =   48
      Top             =   2520
      Width           =   2535
      Begin VB.CommandButton cmd_KeyFobButtons 
         Caption         =   "Find Car"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   54
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmd_KeyFobButtons 
         Caption         =   "PANIC"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   38
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmd_KeyFobButtons 
         Caption         =   "Unlock All"
         Enabled         =   0   'False
         Height          =   735
         Index           =   3
         Left            =   1200
         TabIndex        =   37
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd_KeyFobButtons 
         Caption         =   "Trunk"
         Enabled         =   0   'False
         Height          =   735
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd_KeyFobButtons 
         Caption         =   "Unlock Driver"
         Enabled         =   0   'False
         Height          =   735
         Index           =   1
         Left            =   1200
         TabIndex        =   35
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmd_KeyFobButtons 
         Caption         =   "Lock"
         Enabled         =   0   'False
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cmb_KeyFobVehicle 
         Height          =   315
         ItemData        =   "frm_Main.frx":0000
         Left            =   240
         List            =   "frm_Main.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frm_RadioPhone 
      Caption         =   "Radio Phone Mode"
      Height          =   2895
      Left            =   2880
      TabIndex        =   47
      Top             =   2520
      Width           =   2175
      Begin VB.Timer timer_RadioPhone 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   840
         Top             =   720
      End
      Begin VB.ComboBox cmb_RadioPhoneVehicle 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Main.frx":0004
         Left            =   120
         List            =   "frm_Main.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmd_RadioPhoneStop 
         Caption         =   "Stop Phone"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmd_RadioPhoneStart 
         Caption         =   "Initiate Phone"
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame frm_WheelControls 
      Caption         =   "Wheel Controls"
      Height          =   2895
      Left            =   120
      TabIndex        =   46
      Top             =   2520
      Width           =   2655
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Phone"
         Enabled         =   0   'False
         Height          =   375
         Index           =   9
         Left            =   840
         TabIndex        =   29
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Seek Down"
         Enabled         =   0   'False
         Height          =   495
         Index           =   8
         Left            =   1680
         TabIndex        =   28
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Volume Down"
         Enabled         =   0   'False
         Height          =   495
         Index           =   7
         Left            =   960
         TabIndex        =   27
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Source"
         Enabled         =   0   'False
         Height          =   495
         Index           =   6
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Right"
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   1680
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Mute"
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   960
         TabIndex        =   24
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Left"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Seek Up"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1680
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Volume Up"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   21
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmd_WheelButtons 
         Caption         =   "Info"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cmb_WheelVehicle 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Main.frx":0008
         Left            =   360
         List            =   "frm_Main.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frm_Connection 
      Caption         =   "Connection"
      Height          =   2295
      Left            =   6720
      TabIndex        =   44
      Top             =   120
      Width           =   3135
      Begin MSCommLib.MSComm comm_Serial 
         Left            =   1200
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   256
         NullDiscard     =   -1  'True
         OutBufferSize   =   1024
         ParityReplace   =   0
         RThreshold      =   1
      End
      Begin VB.ComboBox cmb_ConnectionSerialBaud 
         Height          =   315
         ItemData        =   "frm_Main.frx":000C
         Left            =   2040
         List            =   "frm_Main.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmb_ConnectionSerialPort 
         Height          =   315
         ItemData        =   "frm_Main.frx":0080
         Left            =   1080
         List            =   "frm_Main.frx":009C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt_ConnectionType 
         Caption         =   "TCP/IP"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton opt_ConnectionType 
         Caption         =   " Serial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmd_Disconnect 
         Caption         =   "Disconnect"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Connect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.Frame frm_SendText 
      Caption         =   "Send DIC Text"
      Height          =   2295
      Left            =   2880
      TabIndex        =   43
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox cmb_DICTextVehicle 
         Height          =   315
         ItemData        =   "frm_Main.frx":00D0
         Left            =   120
         List            =   "frm_Main.frx":00D2
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmd_DICTextSend 
         Caption         =   "Send Text (F3)"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_DICText 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Text            =   "Welcome to the GMLAN!"
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Frame frm_Chime 
      Caption         =   "Chime"
      Height          =   2295
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   2655
      Begin MSComCtl2.UpDown updown_ChimeDuty 
         Height          =   420
         Left            =   2296
         TabIndex        =   14
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   741
         _Version        =   393216
         Value           =   255
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt_ChimeDuty"
         BuddyDispid     =   196635
         OrigLeft        =   2280
         OrigTop         =   1320
         OrigRight       =   2535
         OrigBottom      =   1695
         Max             =   255
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.TextBox txt_ChimeDuty 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "255"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt_ChimeCount 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "5"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txt_ChimeDelay 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "120"
         Top             =   1320
         Width           =   495
      End
      Begin MSComCtl2.UpDown updown_ChimeCount 
         Height          =   420
         Left            =   1456
         TabIndex        =   12
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   741
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txt_ChimeCount"
         BuddyDispid     =   196636
         OrigLeft        =   1440
         OrigTop         =   1320
         OrigRight       =   1695
         OrigBottom      =   1695
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.ComboBox cmb_ChimeType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Main.frx":00D4
         Left            =   120
         List            =   "frm_Main.frx":00D6
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin MSComCtl2.UpDown updown_ChimeDelay 
         Height          =   420
         Left            =   616
         TabIndex        =   10
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   741
         _Version        =   393216
         Value           =   120
         BuddyControl    =   "txt_ChimeDelay"
         BuddyDispid     =   196637
         OrigLeft        =   600
         OrigTop         =   1320
         OrigRight       =   855
         OrigBottom      =   1695
         Max             =   255
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.ComboBox cmb_ChimeVehicle 
         Height          =   315
         ItemData        =   "frm_Main.frx":00D8
         Left            =   120
         List            =   "frm_Main.frx":00DA
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmd_ChimeSend 
         Caption         =   "Send Chime (F2)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lbl_ChimeDuty 
         Caption         =   "Duty:"
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbl_ChimeCount 
         Caption         =   "Count:"
         Height          =   255
         Left            =   960
         TabIndex        =   42
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbl_ChimeDelay 
         Caption         =   "Delay:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2249
            MinWidth        =   2258
            Text            =   "Disconnected"
            TextSave        =   "Disconnected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   132292
            MinWidth        =   132292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SerialComm As MSComm
Public ChimePriority As String
Public ChimeHeader As String
Public DICTextPriorityImpala As String
Public DICTextHeaderImpala As String
Public WheelPriorityImpala As String
Public WheelHeaderImpala As String
Public RadioPhonePriorityImpala As String
Public RadioPhoneHeaderImpala As String
Public KeyFobPriorityImpala As String
Public KeyFobHeaderImpala As String
Public isConnected As Boolean
Public lineEnding As String
Public CTRL_1 As Boolean
Public CTRL_2 As Boolean


'Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Wait(ByVal DurationMS As Long)
    Dim EndTime As Long
    
    EndTime = GetTickCount + DurationMS
    Do While EndTime > GetTickCount
        'DoEvents
        Sleep 1
    Loop
End Sub

Private Sub cmd_RadioPhoneStart_Click()
cmd_RadioPhoneStop.Enabled = True
cmd_RadioPhoneStart.Enabled = False
'open the phone
mscomm_SendHeader RadioPhonePriorityImpala, RadioPhoneHeaderImpala, "Radio Phone:"
mscomm_SendCommand "0E00", "Radio Phone:"
mscomm_SendCommand "0E11", "Radio Phone:"
mscomm_SendCommand "0E14", "Radio Phone:"
'the timer keeps sending the keepalive data
timer_RadioPhone.Enabled = True
End Sub

Private Sub cmd_RadioPhoneStop_Click()
cmd_RadioPhoneStop.Enabled = False
cmd_RadioPhoneStart.Enabled = True
timer_RadioPhone.Enabled = False
'close the connection gracefully...
mscomm_SendCommand "0E12", "Radio Phone:"
mscomm_SendCommand "0E10", "Radio Phone:"
mscomm_SendCommand "0E00", "Radio Phone:"
End Sub

Private Sub comm_Serial_OnComm()
'MsgBox "a"
'
' Assumes that MSComm1.RThreshold = 1
'
Dim strData As String
Dim strItems() As String
Static strBuffer As String
Dim boComplete As Boolean
Select Case comm_Serial.CommEvent
    Case comEvReceive
    '
    ' Something's ready to be received
    ' Read it and append it to the buffer
    '
        strData = comm_Serial.Input
        strBuffer = strBuffer & strData
        'Do
            txt_Buffer.Text = txt_Buffer.Text & strBuffer
            txt_Buffer.SelStart = Len(txt_Buffer.Text)
                    '
                    ' All characters in the buffer have been processed
                    ' flush it and singnal to exit the loop
                    '
                    strBuffer = ""
                    boComplete = True
End Select
End Sub

Public Sub setupCommands()
lineEnding = Chr(13) ' & Chr(10) '0d +0a
ChimePriority = "10"
ChimeHeader = "01E058"
DICTextPriorityImpala = "10"
DICTextHeaderImpala = "306060"
WheelPriorityImpala = "10"
WheelHeaderImpala = "0D0040"
RadioPhonePriorityImpala = "10"
RadioPhoneHeaderImpala = "31C097"
KeyFobPriorityImpala = "08"
KeyFobHeaderImpala = "0080B0"
End Sub

Public Sub mscomm_SendCommand(CommandData As String, Name As String)
If (isConnected) Then
Dim pnlStatus As Panel
Set pnlStatus = StatusBar1.Panels.Item(2)
'Only set this when mscomm is actually reporting that it is disconnected to the serial port
'or Winsock is disconnected
pnlStatus.Text = "Last command - " & Name & " " & CommandData

'MsgBox CommandData, vbInformation, "SendCommand"
'use the public instance of the mscomm

'send the command...
comm_Serial.Output = CommandData & lineEnding
Wait 30

End If
End Sub

Public Sub mscomm_SendHeader(Priority As String, Header As String, Name As String)
Dim DataToSendPriority As String
Dim DataToSendHeader As String
If (isConnected) Then
Dim pnlStatus As Panel
Set pnlStatus = StatusBar1.Panels.Item(2)
'Only set this when mscomm is actually reporting that it is disconnected to the serial port
'or Winsock is disconnected
pnlStatus.Text = "Last command - " & Name & " " & Priority & " " & Header


DataToSendPriority = " ATCP " & Priority & lineEnding
DataToSendHeader = " ATSH " & Header & lineEnding

comm_Serial.Output = DataToSendPriority
Wait 80
comm_Serial.Output = DataToSendHeader
Wait 80

'MsgBox DataToSendHeader, vbInformation, DataToSendPriority
End If 'Connected
End Sub

Private Sub cmd_ChimeSend_Click()
'Actual Chime Command:
'mscomm_SendCommand "86 78 05 FF 05"
ChimeDelay = Right("00" & Hex(txt_ChimeDelay.Text), 2)
ChimeCount = Right("00" & Hex(txt_ChimeCount.Text), 2)
ChimeDuty = Right("00" & Hex(txt_ChimeDuty.Text), 2)
ChimeType = cmb_ChimeType.ListIndex
Dim ChimeCommand As String
ChimeCommand = "8" & ChimeType & " " & ChimeDelay & " " & ChimeCount & " " & ChimeDuty & " 05"
'mscomm_SendCommand " ", "Chime:"
mscomm_SendHeader ChimePriority, ChimeHeader, "Chime Header:"
'Wait 80
mscomm_SendCommand ChimeCommand, "Chime:"
End Sub

Private Sub cmd_Connect_Click()
comm_Serial.CommPort = cmb_ConnectionSerialPort.ItemData(cmb_ConnectionSerialPort.ListIndex)
comm_Serial.Settings = cmb_ConnectionSerialBaud.ItemData(cmb_ConnectionSerialBaud.ListIndex) & ",n,8,1"
comm_Serial.PortOpen = True

If (comm_Serial.PortOpen) Then
'Set up for 29bit GMLAN/SWCAN!

comm_Serial.Output = "ATWS" & lineEnding
Wait 200
comm_Serial.Output = "ATL1" & lineEnding
Wait 30
comm_Serial.Output = "ATPP 2D SV 0F" & lineEnding
Wait 30
comm_Serial.Output = "ATPP 2C SV 40" & lineEnding
Wait 30
comm_Serial.Output = "ATPP 2D ON" & lineEnding
Wait 30
comm_Serial.Output = "ATPP 2C ON" & lineEnding
Wait 30
comm_Serial.Output = "ATPP 2A OFF" & lineEnding
Wait 30
comm_Serial.Output = "ATWS" & lineEnding
Wait 200
comm_Serial.Output = "ATL1" & lineEnding
Wait 30
comm_Serial.Output = "ATCAF1" & lineEnding
Wait 30
comm_Serial.Output = "ATSPB" & lineEnding
Wait 30
comm_Serial.Output = "ATH1" & lineEnding
Wait 30
comm_Serial.Output = "ATR0" & lineEnding
Wait 30
comm_Serial.Output = "ATL1" & lineEnding
Wait 40

'StatusBar1.Panels.Item.Text
Dim pnlConnected As Panel
Set pnlConnected = StatusBar1.Panels.Item(1)
'Only set this when mscomm is actually reporting that it is connected to the serial port
'or Winsock is connected
pnlConnected.Text = "Connected"
isConnected = True

cmd_Connect.Enabled = False
cmd_Disconnect.Enabled = True
cmd_ChimeSend.Enabled = True
cmb_ChimeVehicle.Enabled = True
cmb_ChimeType.Enabled = True
cmd_DICTextSend.Enabled = False
cmb_DICTextVehicle.Enabled = True
txt_DICText.Enabled = True
cmd_RadioPhoneStart.Enabled = False
cmb_RadioPhoneVehicle.Enabled = True

cmd_WheelButtons(0).Enabled = True
cmd_WheelButtons(1).Enabled = True
cmd_WheelButtons(2).Enabled = True
cmd_WheelButtons(3).Enabled = True
cmd_WheelButtons(4).Enabled = True
cmd_WheelButtons(5).Enabled = True
cmd_WheelButtons(6).Enabled = True
cmd_WheelButtons(7).Enabled = True
cmd_WheelButtons(8).Enabled = True
cmd_WheelButtons(9).Enabled = True

cmb_WheelVehicle.Enabled = True

cmd_KeyFobButtons(0).Enabled = True
cmd_KeyFobButtons(1).Enabled = True
cmd_KeyFobButtons(2).Enabled = True
cmd_KeyFobButtons(3).Enabled = True
cmd_KeyFobButtons(4).Enabled = True
cmd_KeyFobButtons(5).Enabled = True

cmb_KeyFobVehicle.Enabled = True

txt_ChimeDelay.Enabled = True
txt_ChimeCount.Enabled = True
txt_ChimeDuty.Enabled = True
updown_ChimeDelay.Enabled = True
updown_ChimeCount.Enabled = True
updown_ChimeDuty.Enabled = True

WhichOneConnected = 0
If opt_ConnectionType(0).Value Then
WhichOneConnected = 0
End If
If opt_ConnectionType(1).Value Then
WhichOneConnected = 1
End If

opt_ConnectionType(0).Enabled = False
opt_ConnectionType(1).Enabled = False

cmb_ConnectionSerialPort.Enabled = False
cmb_ConnectionSerialBaud.Enabled = False


opt_ConnectionType(WhichOneConnected).Value = True
End If

End Sub

Private Sub cmd_DICTextSend_Click()
'IF CONNECTED
'call send command sub
'mscomm_SendCommand DICTextArb
'mscomm_SendCommand DICTextHeader
'mscomm_SendCommand "00 00 00 00" + whatever
End Sub

Private Sub cmd_Disconnect_Click()
If (comm_Serial.PortOpen) Then
comm_Serial.PortOpen = False
End If
If (Not comm_Serial.PortOpen) Then
Dim pnlConnected As Panel
Set pnlConnected = StatusBar1.Panels.Item(1)
'Only set this when mscomm is actually reporting that it is disconnected to the serial port
'or Winsock is disconnected
pnlConnected.Text = "Disconnected"
Dim pnlStatus As Panel
Set pnlStatus = StatusBar1.Panels.Item(2)
'Only set this when mscomm is actually reporting that it is disconnected to the serial port
'or Winsock is disconnected
pnlStatus.Text = ""
isConnected = False

cmd_Connect.Enabled = True
cmd_Disconnect.Enabled = False
cmd_ChimeSend.Enabled = False
cmb_ChimeType.Enabled = False
cmb_ChimeVehicle.Enabled = False
cmd_DICTextSend.Enabled = False
cmb_DICTextVehicle.Enabled = False
txt_DICText.Enabled = False
cmd_RadioPhoneStart.Enabled = False
cmb_RadioPhoneVehicle.Enabled = False
cmb_ConnectionSerialPort.Enabled = True
cmb_ConnectionSerialBaud.Enabled = True

cmd_WheelButtons(0).Enabled = False
cmd_WheelButtons(1).Enabled = False
cmd_WheelButtons(2).Enabled = False
cmd_WheelButtons(3).Enabled = False
cmd_WheelButtons(4).Enabled = False
cmd_WheelButtons(5).Enabled = False
cmd_WheelButtons(6).Enabled = False
cmd_WheelButtons(7).Enabled = False
cmd_WheelButtons(8).Enabled = False
cmd_WheelButtons(9).Enabled = False

cmb_WheelVehicle.Enabled = False

cmd_KeyFobButtons(0).Enabled = False
cmd_KeyFobButtons(1).Enabled = False
cmd_KeyFobButtons(2).Enabled = False
cmd_KeyFobButtons(3).Enabled = False
cmd_KeyFobButtons(4).Enabled = False
cmd_KeyFobButtons(5).Enabled = False

cmb_KeyFobVehicle.Enabled = False

opt_ConnectionType(0).Enabled = True
opt_ConnectionType(1).Enabled = True

txt_ChimeDelay.Enabled = False
txt_ChimeCount.Enabled = False
txt_ChimeDuty.Enabled = False
updown_ChimeDelay.Enabled = False
updown_ChimeCount.Enabled = False
updown_ChimeDuty.Enabled = False

End If
End Sub

Private Sub cmd_KeyFobButtons_Click(Index As Integer)
'
Dim KeyFobCommand As String
Dim KeyFobCommandName As String
Select Case Index
    Case 0 'Lock
    KeyFobCommand = "02 01"
    KeyFobCommandName = "Lock"
    Case 1 'Unlock Driver
    KeyFobCommand = "02 02"
    KeyFobCommandName = "Unlock Driver"
    Case 2 'Trunk
    KeyFobCommand = "02 04"
    KeyFobCommandName = "Trunk"
    Case 3 'Unlock All
    KeyFobCommand = "02 03"
    KeyFobCommandName = "Unlock All"
    Case 4 'Panic
    KeyFobCommand = "02 07" '07 alarm '0E find car beeps
    KeyFobCommandName = "Panic"
    Case 5 'Find Car
    KeyFobCommand = "02 0E" '07 alarm '0E find car beeps
    KeyFobCommandName = "Find Car"
    Case Else 'Do nothing
    KeyFobCommand = ""
    KeyFobCommandName = "Invalid Button"
End Select

mscomm_SendHeader KeyFobPriorityImpala, KeyFobHeaderImpala, "KeyFob Header:"
mscomm_SendCommand KeyFobCommand, "KeyFob Button " & KeyFobCommandName & ":"
End Sub

Private Sub cmd_ShowLog_Click()
'
Load dlg_Log
dlg_Log.Visible = True
End Sub

Private Sub cmd_WheelButtons_Click(Index As Integer)
'Click that button
Dim WheelButtonCommand As String
Dim WheelButtonCommandName As String
Dim WheelButtonStopCommand As String
WheelButtonStopCommand = "00"
Select Case Index
    Case 0 'Info
    WheelButtonCommand = "0B"
    WheelButtonCommandName = "Info"
    Case 1 'Volume Up
    WheelButtonCommand = "03"
    WheelButtonCommandName = "Volume Up"
    Case 2 'Seek Up
    WheelButtonCommand = "07"
    WheelButtonCommandName = "Seek Up"
    Case 3 'Left
    WheelButtonCommand = "05"
    WheelButtonCommandName = "Left"
    Case 4 'Mute
    WheelButtonCommand = "01"
    WheelButtonCommandName = "Mute"
    Case 5 'Right
    WheelButtonCommand = "04"
    WheelButtonCommandName = "Right"
    Case 6 'Source
    WheelButtonCommand = "06"
    WheelButtonCommandName = "Source"
    Case 7 'Volume Down
    WheelButtonCommand = "02"
    WheelButtonCommandName = "Volume Down"
    Case 8 'Seek Down
    WheelButtonCommand = "09"
    WheelButtonCommandName = "Seek Down"
    Case 9 'Phone
    WheelButtonCommand = "01" 'Hold This command for 1 sec before 00
    WheelButtonCommandName = "Phone"
    
    Case Else 'Do nothing
    WheelButtonCommand = ""
    WheelButtonCommandName = "Invalid Button"
End Select

If (Index <> 9) Then
mscomm_SendHeader WheelPriorityImpala, WheelHeaderImpala, "Wheel Header:"
mscomm_SendCommand WheelButtonCommand, "Wheel Button " & KeyFobCommandName & ":"
mscomm_SendCommand WheelButtonStopCommand, "Wheel Button " & KeyFobCommandName & ":"
End If

If (Index = 9) Then
mscomm_SendHeader WheelPriorityImpala, WheelHeaderImpala, "Wheel Header:"
mscomm_SendCommand WheelButtonCommand, "Wheel Button " & KeyFobCommandName & ":"
Wait 3000
mscomm_SendCommand WheelButtonStopCommand, "Wheel Button " & KeyFobCommandName & ":"
End If

End Sub

Private Sub cmd_WheelButtons_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'Dim pnlStatus As Panel
'Set pnlStatus = StatusBar1.Panels.Item(2)
'Only set this when mscomm is actually reporting that it is disconnected to the serial port
'or Winsock is disconnected
'pnlStatus.Text = "Index: " + Str(Index) + " KeyDown"
End Sub

Private Sub cmd_WheelButtons_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'Dim pnlStatus As Panel
'Set pnlStatus = StatusBar1.Panels.Item(2)
'Only set this when mscomm is actually reporting that it is disconnected to the serial port
'or Winsock is disconnected
'pnlStatus.Text = "Index: " + Str(Index) + " KeyUp"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If (isConnected) Then
'MsgBox KeyCode
    If KeyCode = vbKeyF2 Then
        cmd_ChimeSend_Click
    End If
    If KeyCode = vbKeyF3 Then
        cmd_DICTextSend_Click
    End If
    If KeyCode = vbKeyW And CTRL_1 Then 'Up - W
        cmd_WheelButtons_Click 1
        KeyCode = 0
    End If
    If KeyCode = vbKeyS And CTRL_1 Then 'Down - S
        cmd_WheelButtons_Click 7
        KeyCode = 0
    End If
    If KeyCode = vbKeyA And CTRL_1 Then 'Left - A
        cmd_WheelButtons_Click 3
        KeyCode = 0
    End If
    If KeyCode = vbKeyD And CTRL_1 Then 'Right - D
        cmd_WheelButtons_Click 5
        KeyCode = 0
    End If
    If KeyCode = vbKeyR And CTRL_1 Then 'Seek Up - R
        cmd_WheelButtons_Click 2
    End If
    If KeyCode = vbKeyF And CTRL_1 Then 'Seek Down - F
        cmd_WheelButtons_Click 8
    End If
    If KeyCode = vbKeyE And CTRL_1 Then 'Info - E
        cmd_WheelButtons_Click 0
    End If
    If KeyCode = vbKeyQ And CTRL_1 Then 'Source - Q
        cmd_WheelButtons_Click 6
    End If
    If KeyCode = vbKeySpace And CTRL_1 Then 'Space - Mute
        cmd_WheelButtons_Click 4
    End If
    If KeyCode = vbKeyZ And CTRL_1 Then 'Phone - Z
        cmd_WheelButtons_Click 9
    End If
    
    If KeyCode = vbKeyI And CTRL_1 Then 'Lock - I
        cmd_KeyFobButtons_Click 0
    End If
    If KeyCode = vbKeyO And CTRL_1 Then 'Unlock Driver - O
        cmd_KeyFobButtons_Click 1
    End If
    If KeyCode = vbKeyK And CTRL_1 Then 'Trunk - K
        cmd_KeyFobButtons_Click 2
    End If
    If KeyCode = vbKeyL And CTRL_1 Then 'Unlock All - L
        cmd_KeyFobButtons_Click 3
    End If
    If KeyCode = vbKeyP And CTRL_1 Then 'Panic - P
        cmd_KeyFobButtons_Click 4
    End If
    If KeyCode = vbKeyU And CTRL_1 Then 'Find Car - U
        cmd_KeyFobButtons_Click 5
    End If
    
End If 'connected
    If KeyCode = vbKeyF10 Then
        If Not CTRL_1 Then
            CTRL_1 = True
            lbl_Shortcuts.Caption = "F10 - Keyboard shortcuts active"
        Else
            CTRL_1 = False
            lbl_Shortcuts.Caption = "F10 - Toggle KB Shortcuts"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If CTRL_1 Then
    KeyAscii = 0
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'CTRL_1 = False
End Sub

Private Sub Form_Load()
setupCommands
'cmb_ChimeVehicle.AddItem "Vehicle", 0
cmb_ChimeVehicle.AddItem "Impala", 0
cmb_ChimeVehicle.AddItem "G8", 1
cmb_ChimeVehicle.Enabled = False
cmb_ChimeVehicle.ListIndex = 0


'cmb_ChimeType.AddItem "Type", 0
cmb_ChimeType.AddItem "Nothing", 0
cmb_ChimeType.AddItem "Blinker - Tick", 1
cmb_ChimeType.AddItem "Blinker - Tock", 2
cmb_ChimeType.AddItem "Blinker - Tick-Tock", 3
cmb_ChimeType.AddItem "Beep", 4
cmb_ChimeType.AddItem "High Beep", 5
cmb_ChimeType.AddItem "Normal Chime", 6
cmb_ChimeType.AddItem "High Chime", 7
cmb_ChimeType.ListIndex = 6

cmb_DICTextVehicle.Enabled = False
cmb_DICTextVehicle.AddItem "Impala", 0
cmb_DICTextVehicle.ListIndex = 0
cmb_WheelVehicle.Enabled = False
cmb_WheelVehicle.AddItem "Impala", 0
cmb_WheelVehicle.ListIndex = 0
cmb_RadioPhoneVehicle.Enabled = False
cmb_RadioPhoneVehicle.AddItem "Impala", 0
cmb_RadioPhoneVehicle.ListIndex = 0
cmb_KeyFobVehicle.Enabled = False
cmb_KeyFobVehicle.AddItem "Impala", 0
cmb_KeyFobVehicle.ListIndex = 0

cmd_RadioPhoneStart.Enabled = False
cmd_ChimeSend.Enabled = False
cmd_DICTextSend.Enabled = False
cmd_Disconnect.Enabled = False

cmb_ConnectionSerialPort.ListIndex = 6
cmb_ConnectionSerialBaud.ListIndex = 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
cmd_Disconnect_Click
End Sub

Private Sub timer_RadioPhone_Timer()
mscomm_SendCommand "0E15", "Radio Phone:"
End Sub
