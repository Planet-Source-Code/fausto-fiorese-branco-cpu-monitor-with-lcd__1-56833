Attribute VB_Name = "modMBMAccess"
Option Explicit

'*********************************************************************************************
'*  API Declarations to open the shared memory and to copy the contained information ...
'*********************************************************************************************
Private Declare Sub CopyMemoryRead Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryWrite Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Const FILE_MAP_WRITE = &H2
Private Const FILE_MAP_READ = &H4

'*********************************************************************************************
'* Generic Defines for MBM ...
'*********************************************************************************************
Public Const MBMnumSensors = 99
Public Const MBMnumVoltages = 10
Public Const MBMnumFans = 10
Public Const MBMnumCPUs = 4

Public Enum MBMBusType
    btISA = 0
    btSMBus = 1
    btVIA686ABus = 2
    btDirectIO = 3
End Enum

Public Enum MBMSMBType
    smtSMBIntel = 0
    smtSMBAMD = 1
    smtSMBALi = 2
    smtSMBNForce = 3
    smtSMBSIS = 4
End Enum

Public Enum MBMSensorType
    stUnknown = 0
    stTemperature = 1
    stVoltage = 2
    stFan = 3
    stMhz = 4
    stPercentage = 5
End Enum

'*********************************************************************************************
'* Shared Data Type Definitions for VB
'*********************************************************************************************

Public Type MBMSharedSensor
      ssType As Byte                                    ': TSensorType;                      // type of sensor
      ssName As String * 12                             ': array [0..11] of AnsiChar;        // name of sensor
      ssPad1 As String * 3                              ': array [0..2] of Char              // padding of 3 byte
      ssCurrent As Double                               ': Double;                           // current value
      ssLow As Double                                   ': Double;                           // lowest readout
      ssHigh As Double                                  ': Double;                           // highest readout
      ssCount As Long                                   ': LongInt;                          // total number of readout
      ssPad2 As String * 4                              ': array [0..3] of Char              // padding of 4 byte
      ssTotal As Double                                 ': Extended;                         // total amout of all readouts
      ssPad3 As String * 6                              ': array [0..5] of Char              // padding of 6 byte
      ssAlarm1 As Double                                ': Double;                           // temp & fan: low alarm; voltage: % off;
      ssAlarm2 As Double                                ': Double;                           // temp: high alarm
End Type
    
Public Type MBMSharedIndex
      iType As MBMSensorType                            ': TSensorType;                          // type of sensor
      Count As Integer                                  ': integer;                              // number of sensor for that type
End Type

Public Type MBMSharedInfo
      siSMB_Base As Integer                             ': Word;                      // SMBus base address
      siSMB_Type As Byte                                ': TBusType;                  // SMBus/Isa bus used to access chip
      siSMB_Code As Byte                                ': TSMBType;                  // SMBus sub type, Intel, AMD or ALi
      siSMB_Addr As Byte                                ': Byte;                      // Address of sensor chip on SMBus
      siSMB_Name As String * 41                         ': array [0..40] of AnsiChar; // Nice name for SMBus
      siISA_Base As Integer                             ': Word;                      // ISA base address of sensor chip on ISA
      siChipType As Long                                ': Integer;                   // Chip nr, connects with Chipinfo.ini
      siVoltageSubType As Byte                          ': Byte;                      // Subvoltage option selected
End Type

Public Type MBMSharedData
      sdVersion As Double                               ': Double;                           // version number (example: 51090)
      sdIndex(0 To 9) As MBMSharedIndex                 ': array [0..9]   of TSharedIndex;   // Sensor index
      sdSensor(0 To 99) As MBMSharedSensor              ': array [0..99]  of TSharedSensor;  // sensor info
      sdInfo As MBMSharedInfo                           ': TSharedInfo;                      // misc. info
      sdStart As String * 41                            ': array [0..40]  of AnsiChar;       // start time
      sdCurrent As String * 41                          ': array [0..40]  of AnsiChar;       // current time
      sdPath As String * 256                            ': array [0..255] of AnsiChar;       // MBM path
End Type

Public Function XTrim(sStr As String) As String

    Dim oStr As String
    oStr = RTrim(LTrim(sStr))
    
    Dim pos As Integer
    Dim l As Integer
    l = Len(oStr)
    If l > 0 Then
    
        For pos = l To 0 Step -1
            If Mid(oStr, l, 1) = Chr(0) Then
                    oStr = Left(oStr, l - 1)
                    l = l - 1
                Else
                    Exit For
            End If
            If l = 0 Then Exit For
        Next pos
    
    End If
    XTrim = oStr

End Function

Public Function MBM_GetSharedData(Optional bSilent As Boolean = True) As MBMSharedData

    Static myDataStruct As MBMSharedData

    Dim myMBMFile As Long
    Dim myMBMMem As Long

    myMBMFile = OpenFileMapping(FILE_MAP_READ, False, "$M$B$M$5$S$D$")
    If myMBMFile = 0 Then
        If (bSilent) Then
            Exit Function
            Else
            MsgBox "MBM Data File/Mem could not be opened. Sorry.": Exit Function
        End If
    End If

    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_READ, 0, 0, 0)
    
    'Fetch the Data Structure
    CopyMemoryRead myDataStruct, myMBMMem, Len(myDataStruct)
    
    'Close Handles
    UnmapViewOfFile myMBMMem
    CloseHandle myMBMFile

    MBM_GetSharedData = myDataStruct

End Function


Public Function MBM_SetSensorValue(SensorID As Integer, Value As Integer) As Integer

    Static myDataStruct As MBMSharedData

    Dim myMBMFile As Long
    Dim myMBMMem As Long
    
    myMBMFile = OpenFileMapping(FILE_MAP_WRITE, False, "$M$B$M$5$S$D$")
    If myMBMFile = 0 Then
        MBM_SetSensorValue = 0
        Exit Function
    End If
    
    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_WRITE, 0, 0, 0)
    
    'Fetch the Data Structure
    CopyMemoryRead myDataStruct, myMBMMem, Len(myDataStruct)

    'Modify the selected Sensor
    myDataStruct.sdSensor(SensorID).ssCurrent = Value

    'Write whole Structure Back
    CopyMemoryWrite myMBMMem, myDataStruct, Len(myDataStruct)

    'Close Handles
    UnmapViewOfFile myMBMMem
    CloseHandle myMBMFile
    
    MBM_SetSensorValue = Value

End Function


