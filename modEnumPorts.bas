Attribute VB_Name = "modEnumPorts"
    '--------------------------------------------------------------------------------
    '    Component  : frmEnumPorts
    '    Project    : prjPorts
    '
    '    Description: Below module use to enumerate all ports on PC
    '
    '    Created    : Ajay Rana on 15th Sep 2004
    '
    '--------------------------------------------------------------------------------
    
    '********************************************************************************
    '   COM_Detect modEnumPorts
    '
    '   Modifications by Mark Mokoski for use in this application
    '
    '   Mark Mokoski
    '   markm@cmtelephone.com
    '   28-JAN-2005
    '********************************************************************************

    Option Explicit

        Private Type PORT_INFO_2
            pPortName                                                                As String
            pMonitorName                                                             As String
            pDescription                                                             As String
            fPortType                                                                As Long
            Reserved                                                                 As Long
        End Type

        Private Type API_PORT_INFO_2
            pPortName                                                                As Long
            pMonitorName                                                             As Long
            pDescription                                                             As Long
            fPortType                                                                As Long
            Reserved                                                                 As Long
        End Type

    Private Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
    Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GetProcessHeap Lib "kernel32" () As Long
    Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
    Dim Ports(0 To 100)                                                              As PORT_INFO_2

Private Function TrimStr(strName As String) As String

    'Finds a null then trims the string

    Dim x            As Integer

    x = InStr(strName, vbNullChar)

        If x > 0 Then TrimStr = Left(strName, x - 1) Else TrimStr = strName

End Function

Private Function LPSTRtoSTRING(ByVal lngPointer As Long) As String

    Dim lngLength            As Long

    'Get number of characters in string
    lngLength = lstrlenW(lngPointer) * 2
    'Initialize string so we have something to copy the string into
    LPSTRtoSTRING = String(lngLength, 0)
    'Copy the string
    CopyMem ByVal StrPtr(LPSTRtoSTRING), ByVal lngPointer, lngLength
    'Convert to Unicode
    LPSTRtoSTRING = TrimStr(StrConv(LPSTRtoSTRING, vbUnicode))

End Function

    'Use ServerName to specify the name of a Remote Workstation i.e. "//WIN95WKST"
    'or leave it blank "" to get the ports of the local Machine

Private Function GetAvailablePorts(ServerName As String) As Long

    Dim ret                              As Long
    Dim PortsStruct(0 To 100)            As API_PORT_INFO_2
    Dim pcbNeeded                        As Long
    Dim pcReturned                       As Long
    Dim TempBuff                         As Long
    Dim i                                As Integer

    'Get the amount of bytes needed to contain the data returned by the API call
    ret = EnumPorts(ServerName, 2, TempBuff, 0, pcbNeeded, pcReturned)
    'Allocate the Buffer
    TempBuff = HeapAlloc(GetProcessHeap(), 0, pcbNeeded)
    ret = EnumPorts(ServerName, 2, TempBuff, pcbNeeded, pcbNeeded, pcReturned)

        If ret Then
            'Convert the returned String Pointer Values to VB String Type
            CopyMem PortsStruct(0), ByVal TempBuff, pcbNeeded

                For i = 0 To pcReturned - 1
                    Ports(i).pDescription = LPSTRtoSTRING(PortsStruct(i).pDescription)
                    Ports(i).pPortName = LPSTRtoSTRING(PortsStruct(i).pPortName)
                    Ports(i).pMonitorName = LPSTRtoSTRING(PortsStruct(i).pMonitorName)
                    Ports(i).fPortType = PortsStruct(i).fPortType
                Next

        End If

    GetAvailablePorts = pcReturned
    'Free the Heap Space allocated for the Buffer

        If TempBuff Then HeapFree GetProcessHeap(), 0, TempBuff

    Erase PortsStruct

    Erase PortsStruct

    Erase PortsStruct

    Erase PortsStruct

    Erase PortsStruct

    Erase PortsStruct

    Erase PortsStruct

    Erase PortsStruct
End Function

Public Sub FindPorts(ListControl As Object)

    'Use, Call FindPorts(frm.control)

    Dim NumPorts                         As Long
    Dim i                                As Integer
    Dim COMname                          As String
    
    'Get the Numbers of Ports in the System
    'and Fill the Ports Structure
    '**Code_Err: Unused variables
    '**Code_Err: COMname

    NumPorts = GetAvailablePorts("")
    'Show the available Ports
    

    'CC_Comment Out (1/27/2005) MJM:
    'Original Code
    '
    'Me.AutoRedraw = True
    '
    '        For i = 0 To NumPorts - 1
    '            Me.Print Ports(i).pPortName
    '        Next

    'End CC_Comment Out

        For i = 0 To NumPorts - 1
            'Mods to list only one type of port (in this case, only "COM" ports)

                If InStr(1, Ports(i).pPortName, "COM") > 0 Then
                    COMname = Mid(Ports(i).pPortName, 1, 3) & " " & _
                    Mid(Ports(i).pPortName, 4, 1)
                    ListControl.AddItem COMname, _
                    (ListControl.ListCount)

                End If

        Next

End Sub
