VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COM Detect and Test"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMouseOver 
      Interval        =   10
      Left            =   6120
      Top             =   6720
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      Caption         =   "COM Ports Avalable"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   29
      Top             =   4920
      Width           =   4455
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear Selection"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         MouseIcon       =   "Form1.frx":1272
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelectAll 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select All Ports"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         MouseIcon       =   "Form1.frx":19BE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1CC8
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.ListBox PortsList 
         CausesValidation=   0   'False
         Height          =   1860
         ItemData        =   "Form1.frx":210A
         Left            =   120
         List            =   "Form1.frx":210C
         MouseIcon       =   "Form1.frx":210E
         MousePointer    =   99  'Custom
         Style           =   1  'Checkbox
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "Test Progress"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   6015
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox InfoText 
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MouseIcon       =   "Form1.frx":2418
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Form1.frx":2722
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label ProgPercent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "COM Port Test Results"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1720
      Width           =   6015
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   7
         Left            =   4680
         MouseIcon       =   "Form1.frx":272D
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2A37
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   6
         Left            =   3120
         MouseIcon       =   "Form1.frx":2E79
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":3183
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   5
         Left            =   1680
         MouseIcon       =   "Form1.frx":35C5
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":38CF
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   4
         Left            =   240
         MouseIcon       =   "Form1.frx":3D11
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":401B
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   3
         Left            =   4680
         MouseIcon       =   "Form1.frx":445D
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":4767
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   2
         Left            =   3120
         MouseIcon       =   "Form1.frx":4BA9
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":4EB3
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   1
         Left            =   1680
         MouseIcon       =   "Form1.frx":52F5
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":55FF
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COM x"
         Height          =   855
         Index           =   0
         Left            =   240
         MouseIcon       =   "Form1.frx":5A41
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":5D4B
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   18
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   17
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   9
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      MouseIcon       =   "Form1.frx":618D
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":6497
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Timer ResponceTimer 
      Interval        =   1000
      Left            =   6120
      Top             =   6240
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6120
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      MouseIcon       =   "Form1.frx":67A1
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":6AAB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Menu mnuCLose 
      Caption         =   "&Close Application"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About COM Detect and Test"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '*******************************************************
    '   COM_Detect frmMain
    '
    '   Application to find and test COM Ports on a PC
    '   Addtional info provided. ID Port as a Modem, and if
    '   port is in use at time of test
    '
    '   Added 28-JAN-2005, App now enumerates installed ports
    '   in the system, and allows for selected port testing
    '   and cleaned up GUI
    '
    '   Mark Mokoski
    '   markm@cmtelephone.com
    '   28-JAN-2005
    '
    '*******************************************************

    Option Explicit
    Dim comlist                       As Integer
    Dim numport                       As Integer
    Dim listindex                     As Integer
    Dim listports                     As Integer
    Dim selports                      As Integer
    Dim buffer                        As String
    Dim myport                        As Integer
    Dim TimeElapsed                   As Integer
    Dim PortInfo(7)                   As String
    Dim Handshake                     As String
    Dim PortSelected                  As Boolean
    Dim Splitter                      As String
    Dim Pos1                          As Integer
    Dim Pos2                          As Integer
    Dim tmp                           As String
    Dim PortSettings(7)               As String
    Dim ExpandedInfoParent            As Object
    
    'Define Balloon Tool Tip objects
    Dim cmdStartTip                   As New clsTooltips
    Dim cmdEndTip                     As New clsTooltips
    Dim Frame1Tip                     As New clsTooltips
    Dim InfoTextTip                   As New clsTooltips
    Dim COMiconTip(7)                 As New clsTooltips
    Dim PortsListTip                  As New clsTooltips



Private Sub CheckPort(PortNum As Integer, Index As Integer)
    
    ResponceTimer.Enabled = False
      
    PrintText "Testing COM" & Trim(Str(PortNum)) & "..."
   
    'Check for port status

        If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
   
    'Start error handling
    On Error GoTo ErrorHandler
    MSComm1.CommPort = PortNum
    MSComm1.Settings = "19200,N,8,1"
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
    On Error GoTo 0
   
    'Send Modem ID request (numeric responce + "OK")
    MSComm1.Output = "ATI1" & Chr$(13)
        
    'Wait for a response for 1 second. If nothing returns then exit

        If WaitForResponse(1) = False Then GoTo NothingReturned
        
    'If "OK" was returned from the "AT" request, we have a Modem
    'Send extended Modem ID request (Text ID of Modem + "OK")
    MSComm1.Output = "ATI4" & Chr$(13)
        
    'Wait for response for 2 seconds. If nothing returns then exit

        If WaitForResponse(1) = False Then GoTo NothingReturned
    
    'If something returned and ID as Modem ("OK" was returned)
    PrintText ParseBuffer(buffer)
    PrintText "COM" & Trim(Str(PortNum)) & " is a modem."
    MSComm1.PortOpen = False
    COMicon(Index).Picture = LoadResPicture(102, vbResIcon)
    COMicon(Index).Caption = "COM" + Str(PortNum)
    COMicon(Index).Visible = True
    PortType(Index).Caption = "Modem"
    PortInfo(Index) = ParseBuffer(buffer)
    PortType(Index).Visible = True
    PortStatus(Index).Caption = "Installed - Idle"
    PortSettings(Index) = UCase(MSComm1.Settings)
    PortStatus(Index).Visible = True
    COMiconTip(Index).Title = "Modem"
    COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & PortInfo(Index) & vbCrLf & "Click for Details"

    Exit Sub
   
    'Just a COM Port with other than a modem (or nothing) attached
NothingReturned:
    PrintText "COM" & Trim(Str(PortNum)) & " is an installed Serial Port"
    MSComm1.PortOpen = False
    COMicon(Index).Picture = LoadResPicture(105, vbResIcon)
    COMicon(Index).Caption = "COM" + Str(PortNum)
    COMicon(Index).Visible = True
    PortType(Index).Caption = "Serial Port"
    PortInfo(Index) = ""
    PortType(Index).Visible = True
    PortStatus(Index).Caption = "Installed - Idle"
    PortSettings(Index) = UCase(MSComm1.Settings)
    PortStatus(Index).Visible = True
    COMiconTip(Index).Title = "Serial Port"
    COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"
    
    Exit Sub


ErrorHandler:
    
    'Error on COM Port, set message, Icon and lablels to reflect error type
    PrintText portError(Err.Number, PortNum, Index)

End Sub

Private Sub cmdClear_Click()

    listports = PortsList.ListCount - 1
    'Clear all the current selections

        For selports = 0 To listports
            PortsList.Selected(selports) = False
        Next selports

    cmdSelectAll.SetFocus

End Sub

Private Sub cmdSelectAll_Click()


    listports = PortsList.ListCount - 1
    'Clear all the current selections

        For selports = 0 To listports
            PortsList.Selected(selports) = False
        Next selports

    'Set selection to all avalable ports

        For selports = 0 To listports
            PortsList.Selected(selports) = True
        Next selports

    cmdStart.SetFocus

End Sub

Private Sub COMicon_Click(Index As Integer)

    Load frmExpandedInfo
    frmExpandedInfo.Top = frmMain.Top + COMicon(Index).Top + COMicon(Index).Height
    frmExpandedInfo.Left = frmMain.Left + COMicon(Index).Left + COMicon(Index).Width
    Set ExpandedInfoParent = COMicon(Index)
    frmExpandedInfo.Visible = True


    'Get some COM Port info and display in text box
    MSComm1.CommPort = Val(Mid$(COMicon(Index).Caption, 5, 1))
    frmExpandedInfo.AutoRedraw = True
    frmExpandedInfo.Print Space$(1) & "COM Port: " & COMicon(Index).Caption
    frmExpandedInfo.Print Space$(1) & "Port Type: " + PortType(Index).Caption
    frmExpandedInfo.Print Space$(1) & "Modem ID: " + PortInfo(Index)
    frmExpandedInfo.Print Space$(1) & "Settings: " + PortSettings(Index)

    
    'Skip if port is in use by another App

        Select Case PortStatus(Index)
            Case "Port In Use"
    
            Case Else
                frmExpandedInfo.Print Space$(1) & "DTR: " + Str(MSComm1.DTREnable)
                frmExpandedInfo.Print Space$(1) & "RTS: " + Str(MSComm1.RTSEnable)


                Select Case MSComm1.Handshaking
                    Case 0
                        Handshake = "NONE"
                    Case 1
                        Handshake = "Xon/Xoff"
                    Case 2
                        Handshake = "RTS"
                    Case 3
                        Handshake = "RTS & Xon/Xoff"
                End Select
        
            frmExpandedInfo.Print Space$(1) & "Handshaking: " + Handshake
            frmExpandedInfo.SetFocus
        End Select

End Sub

Private Sub cmdStart_Click()
    
    'If Details window visible, unload it
    Unload frmExpandedInfo
    
    'Set ProgressBar Value
    ProgressBar.Value = 0
    ProgressBar.Visible = True
    pbForeColor ProgressBar, vbRed
    
    ProgPercent.Caption = "0 %"
    ProgPercent.Visible = True
        
    PortSelected = False

    comlist = 0
    selports = 0
    
    'See if we have ANY ports selected for test

        Do Until comlist = (PortsList.ListCount)

                If PortSelected = False _
            And PortsList.Selected(comlist) = True _
            Then PortSelected = True

                If PortsList.Selected(comlist) = True Then selports = selports + 1
            comlist = comlist + 1
        Loop
        
    'If no ports selected, fire off message box

        If PortSelected = False Then
            MsgBox "Please select COM Port(s) to test.", vbInformation
            Exit Sub
        End If

    Frame1Tip.Active = True
    
    'change enabled status of some controls and "Grey Out"
    cmdStart.Enabled = False
    cmdStart.BackColor = vbButtonFace
    cmdEnd.Enabled = False
    cmdEnd.BackColor = vbButtonFace
    cmdSelectAll.Enabled = False
    cmdSelectAll.BackColor = vbButtonFace
    cmdClear.Enabled = False
    cmdClear.BackColor = vbButtonFace
    PortsList.Enabled = False
    
    'Hide COM port buttons from privious test

        For comlist = 0 To 7
            COMicon(comlist).Visible = False
            PortType(comlist).Visible = False
            PortStatus(comlist).Visible = False
        Next comlist

    InfoText.Text = ""
        
    'Set COMicon start index
    comlist = 0

        For listindex = 0 To (PortsList.ListCount - 1)
                
            'If list item selected, test the port

                If PortsList.Selected(listindex) = True Then
                    'Get COM Port number from selected list item
                    numport = Val(Mid$(PortsList.List(listindex), 5, 1))
                    Call CheckPort(numport, comlist)
                    'Set Progress Bar and Percent Label
                    ProgressBar.Value = Int(((comlist + 1) / selports) * 100)
                    ProgPercent.Caption = Str(ProgressBar.Value) & " %"
                    'Set the COMicon index for the next test
                    comlist = comlist + 1
                End If

        Next listindex

    'Testing is over, so enable controls for retesting
    
    pbForeColor ProgressBar, vbGreen
    ProgPercent.Caption = "Test Complete"
    cmdStart.Enabled = True
    cmdStart.BackColor = &HC0C0C0
    cmdEnd.Enabled = True
    cmdEnd.BackColor = &HC0C0C0
    cmdSelectAll.Enabled = True
    cmdSelectAll.BackColor = &HC0C0C0
    cmdClear.Enabled = True
    cmdClear.BackColor = &HC0C0C0
    PortsList.Enabled = True
    
    PrintText "COM Port testing complete"
    PrintText "Click START to retest COM Ports"

    Frame1Tip.Active = False
    
End Sub

Private Sub cmdEnd_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Create Balloon Tool Tips
    cmdStartTip.CreateBalloon cmdStart, _
    "Click to Start Testing", _
    "Start COM Test", 1
    
    cmdEndTip.CreateBalloon cmdEnd, _
    "Click to end Program", _
    "End Application", 3
    
    InfoTextTip.CreateBalloon InfoText, _
    "Test Status Messages and COM Port info", _
    "Status Messages", 1
    
    Frame1Tip.CreateBalloon Frame1, _
    "COM Port Status Shown Here", _
    "COM Port Test Results", 1
    
    PortsListTip.CreateBalloon PortsList, _
    "List of COM Ports Detected in System." & vbCrLf & _
    "Check Ports to Test, or click ""Select All Ports"".", _
    "Avalable COM Ports", 1


    'Make Balloon Tool Tips for result command buttons

        For comlist = 0 To 7
            COMiconTip(comlist).CreateBalloon COMicon(comlist), _
            "Serial Port Number" & vbCrLf & _
            "Serial Port Info", _
            "Serial Port", 1

        Next comlist
    
    ResponceTimer.Enabled = False

    'Set Progress text
    InfoText.Text = ""
    PrintText "Click START to begin COM Port Test"
    ProgPercent.Caption = "Click START to begin COM Port Test"
    'Set ProgressBar Value
    ProgressBar.Value = 0
    
    'Go and find installed ports on system
    Call FindPorts(frmMain.PortsList)
    
    'Change Progress bar colors
    pbForeColor ProgressBar, vbRed
    pbBackColor ProgressBar, &H808000
    
End Sub

Private Sub Form_Terminate()

    Unload frmExpandedInfo
    Unload frmAbout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmExpandedInfo
    Unload frmAbout
    
End Sub

Private Sub mnuAbout_Click()

    frmAbout.Visible = True

End Sub

Private Sub mnuCLose_Click()

    Unload Me

End Sub

Private Sub PortsList_Click()

    cmdStart.SetFocus

End Sub

Private Sub ResponceTimer_Timer()

    'Increment the time elaped value (1 Sec)
    TimeElapsed = TimeElapsed + 1

End Sub

Private Function WaitForResponse(Wait_Sec As Integer) As Boolean
    
    buffer = ""
    WaitForResponse = False
    TimeElapsed = 0
    ResponceTimer.Enabled = True
   
        Do
      
            DoEvents
      
            buffer = buffer & MSComm1.Input
      
                If Len(buffer) <> 0 Then

                        If InStr(1, buffer, "OK") <> 0 Then
            
                            WaitForResponse = True
                            ResponceTimer.Enabled = False
                            TimeElapsed = 0
                            
                            Exit Function
                        End If

                End If
   
                If TimeElapsed > Wait_Sec Then
                    ResponceTimer.Enabled = False
                    Exit Function
                End If
   
        Loop
   
End Function

Private Sub PrintText(outText As String)

    InfoText.Text = InfoText.Text + outText + vbCrLf
    InfoText.SelStart = Len(InfoText.Text) + 1

End Sub

Private Function ParseBuffer(readBuf As String) As String
    
    'Gets rid of stuff from the "ADI4" ID request (leading space(s),vbcrlf's and "OK")
    Splitter = Chr(13) & Chr(10)
    Pos1 = InStr(1, readBuf, Splitter)
    Pos2 = InStr(Pos1 + 2, readBuf, Splitter)

    ParseBuffer = Mid(readBuf, Pos1 + 2, Pos2 - Pos1 - 2)

End Function

Private Function portError(ERROR As Long, PortNum As Integer, Index As Integer) As String


        Select Case ERROR
            Case 8021
                tmp = "Internal error retrieving device control block for the port"
                COMicon(Index).Picture = LoadResPicture(104, vbResIcon)
                COMicon(Index).Caption = "COM" + Str(PortNum)
                COMicon(Index).Visible = True
                PortType(Index).Caption = "Serial Port"
                PortInfo(Index) = ""
                PortType(Index).Visible = True
                PortStatus(Index).Caption = tmp
                PortStatus(Index).Visible = True
                COMiconTip(Index).Title = "Serial Port"
                COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"

            Case 394
                tmp = "Property is write-only"
            Case 380
                tmp = "Invalid property value"
            Case 8012
                'tmp = "The device is not open"
                tmp = "Port In Use"
                COMicon(Index).Picture = LoadResPicture(103, vbResIcon)
                COMicon(Index).Caption = "COM" + Str(PortNum)
                COMicon(Index).Visible = True
                PortType(Index).Caption = "Serial Port"
                PortInfo(Index) = ""
                PortType(Index).Visible = True
                PortStatus(Index).Caption = tmp
                PortStatus(Index).Visible = True
                COMiconTip(Index).Title = "Serial Port"
                COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"

            Case 8005
                tmp = "Port In Use"
                COMicon(Index).Picture = LoadResPicture(103, vbResIcon)
                COMicon(Index).Caption = "COM" + Str(PortNum)
                COMicon(Index).Visible = True
                PortType(Index).Caption = "Serial Port"
                PortInfo(Index) = ""
                PortType(Index).Visible = True
                PortStatus(Index).Caption = tmp
                PortStatus(Index).Visible = True
                COMiconTip(Index).Title = "Serial Port"
                COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"

            Case 8002
                tmp = "Invalid port number"
            Case 8018
                tmp = "Operation valid only when the port is open"
                COMicon(Index).Picture = LoadResPicture(104, vbResIcon)
                COMicon(Index).Caption = "COM" + Str(PortNum)
                COMicon(Index).Visible = True
                PortType(Index).Caption = "Serial Port"
                PortInfo(Index) = ""
                PortType(Index).Visible = True
                PortStatus(Index).Caption = tmp
                PortStatus(Index).Visible = True
                COMiconTip(Index).Title = "Serial Port"
                COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"

            Case 8000
                tmp = "Operation not valid while the port is opened"
                COMicon(Index).Picture = LoadResPicture(104, vbResIcon)
                COMicon(Index).Caption = "COM" + Str(PortNum)
                COMicon(Index).Visible = True
                PortType(Index).Caption = "Serial Port"
                PortInfo(Index) = ""
                PortType(Index).Visible = True
                PortStatus(Index).Caption = tmp
                PortStatus(Index).Visible = True
                COMiconTip(Index).Title = "Serial Port"
                COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"

            Case 8020 & 8015
                'tmp = "Error reading comm device"
                tmp = "Port ERROR"
                COMicon(Index).Picture = LoadResPicture(104, vbResIcon)
                COMicon(Index).Caption = "COM" + Str(PortNum)
                COMicon(Index).Visible = True
                PortType(Index).Caption = "Serial Port"
                PortInfo(Index) = ""
                PortType(Index).Visible = True
                PortStatus(Index).Caption = tmp
                PortStatus(Index).Visible = True
                COMiconTip(Index).Title = "Serial Port"
                COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"

            Case 383
                tmp = "Property is read-only"
            Case Else
                tmp = "Other error..."
                COMicon(Index).Picture = LoadResPicture(104, vbResIcon)
                COMicon(Index).Caption = "COM" + Str(PortNum)
                COMicon(Index).Visible = True
                PortType(Index).Caption = "Serial Port"
                PortInfo(Index) = ""
                PortType(Index).Visible = True
                PortStatus(Index).Caption = tmp
                PortStatus(Index).Visible = True
                COMiconTip(Index).Title = "Serial Port"
                COMiconTip(Index).TipText = "COM" & Str(PortNum) & " Installed" & vbCrLf & "Click for Details"

        End Select

    portError = tmp
    pbForeColor ProgressBar, vbYellow
    
End Function

Private Sub tmrMouseOver_Timer()

    '"Mouse Over" check
    'Yes I know it uses a timer, but the code works.
    'Subclassing and checking "WM_" messages
    'just for this one control array seemed a bit over kill
    'for closing a tool form
    '
    'See modMouseOver for code details
    '
    '
    'Sample
    '    On Error Resume Next
    '
    '    If IsMouseOver(conrtol) Then
    '        do something
    '    Else
    '        do something else
    '    End If
    
    On Error Resume Next

    'Control that opened the frmExpandedInfo window is stored
    'in the ExpandedInfoParent object.
    'On "MouseLeave", close (unload) frmExpandedInfo window

        If frmExpandedInfo.Visible = True Then

                If Not IsMouseOver(ExpandedInfoParent) Then
                    Unload frmExpandedInfo
                End If

        End If

End Sub
