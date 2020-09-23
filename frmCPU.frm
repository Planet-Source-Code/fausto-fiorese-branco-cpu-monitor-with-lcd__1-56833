VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCPU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CPU Monitor"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3555
      TabIndex        =   15
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3420
      Top             =   1125
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3420
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      OutBufferSize   =   1024
      BaudRate        =   19200
   End
   Begin VB.Label lblCooler 
      Height          =   240
      Left            =   1260
      TabIndex        =   14
      Top             =   2025
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Cooler:"
      Height          =   240
      Index           =   6
      Left            =   90
      TabIndex        =   13
      Top             =   2025
      Width           =   1185
   End
   Begin VB.Label lblUpTime 
      Height          =   240
      Left            =   2970
      TabIndex        =   12
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblVolts2 
      Height          =   240
      Left            =   1260
      TabIndex        =   11
      Top             =   1710
      Width           =   1815
   End
   Begin VB.Label lblVolts1 
      Height          =   240
      Left            =   1260
      TabIndex        =   10
      Top             =   1395
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Volts:"
      Height          =   240
      Index           =   5
      Left            =   90
      TabIndex        =   9
      Top             =   1395
      Width           =   1185
   End
   Begin VB.Label lblTemperatura 
      Height          =   240
      Left            =   1260
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Temps:"
      Height          =   240
      Index           =   4
      Left            =   90
      TabIndex        =   7
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label lblHD1 
      Height          =   240
      Left            =   1260
      TabIndex        =   6
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "HD1:"
      Height          =   240
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   765
      Width           =   1185
   End
   Begin VB.Label lblMemoria 
      Height          =   240
      Left            =   1260
      TabIndex        =   4
      Top             =   450
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Memory:"
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   450
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Up-Time:"
      Height          =   240
      Index           =   1
      Left            =   2115
      TabIndex        =   2
      Top             =   135
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Processador:"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   1185
   End
   Begin VB.Label lblProcessador 
      Caption         =   "0%"
      Height          =   240
      Left            =   1305
      TabIndex        =   0
      Top             =   135
      Width           =   690
   End
End
Attribute VB_Name = "frmCPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const control_char = "&h5C"
Const byte127_char = "&h5F"
Const font_char = "&h40"
Const ignore_char = "&h5D"
Const display_char = "&h5B"

Private QueryObject As Object
Private Function Dados_Sensores(Sensor As Integer) As String

    Dim infoStr As String
    Dim valueStr As String
    
    Dim myData As MBMSharedData
    myData = MBM_GetSharedData
    
    Dados_Sensores = myData.sdSensor(Sensor).ssCurrent

End Function

Private Sub cmdSair_Click()

    End

End Sub

Private Sub Form_Load()
    
   SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
   SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
   
   If IsWinNT Then
       Set QueryObject = New clsCPUUsageNT
   Else
       Set QueryObject = New clsCPUUsage
   End If
   
   QueryObject.Initialize
   
   If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
   End If
   MSComm1.CommPort = 1
   MSComm1.PortOpen = True

   send_to_lcd control_char & "&h40&h" & Hex(32) & "&h30"
End Sub


Private Sub send_to_lcd(inputx As String, Optional centerit As Boolean = False)
   Dim outputx As String
   Dim i As Long
   
   outputx = ""
   For i = 0 To Len(inputx)
      ' if next two characters are "&h" convert to Hex value
      If LCase(Mid(inputx, i + 1, 2)) = "&h" Then
         outputx = outputx & Chr(Val(Mid(inputx, i + 1, 4)))
         If i + 2 <= Len(inputx) Then
            i = i + 3
         Else
            i = Len(inputx)
         End If
      Else
         ' need to send two "\" characters to make one show on LCD
         If Mid(inputx, i + 1, 1) = "\" Then
            outputx = outputx & "\\"
         Else
            outputx = outputx & Mid(inputx, i + 1, 1)
         End If
      End If
   Next
   ' center the data on the LCD
   If centerit Then
      If Len(outputx) > 24 Then
         outputx = Left(outputx, 24)
      End If
      outputx = Left("                        ", Int((24 - Len(outputx)) / 2)) & outputx
   End If
   ' send character string to Comm port
   MSComm1.Output = outputx
End Sub

Private Sub Timer1_Timer()
    Dim Ret As Long
    Dim Espaco As Double
    Dim Cont As Integer
    Dim ds_Espaco As String * 2
    Dim Totp1 As Double
    Dim Availp1 As Double
    Dim lUpTime As Long
    Dim upDias As String
    Dim upHoras As String
    Dim upMinutos As String
    Dim upSegundos As String
    
    Ret = QueryObject.Query
    If Ret = -1 Then
        Timer1.Enabled = False
        lblProcessador.Caption = "Error"
        MsgBox "Error while retrieving CPU usage"
    Else
        lblProcessador.Caption = CStr(Ret) + "%"
    End If
    
    GlobalMemoryStatus memoryInfo
    Totp1 = Int(memoryInfo.dwTotalPhys / 1024)
    Availp1 = Int(memoryInfo.dwAvailPhys / 1024)

    Cont = 0
    Espaco = DiscSpace("C:", 0)
    Do While Espaco > 1024
       Cont = Cont + 1
       Espaco = Espaco / 1024
    Loop
    Select Case Cont
           Case 0
           Case 1
              ds_Espaco = "Kb"
           Case 2
              ds_Espaco = "Mb"
           Case 3
              ds_Espaco = "Gb"
    End Select
    Espaco = Format(Espaco, "###0.00")
    
    
    lUpTime = GetTickCount
    
    lUpTime = lUpTime \ 1000
    upSegundos = (lUpTime - ((lUpTime \ 60) * 60))
    lUpTime = lUpTime \ 60
    upMinutos = (lUpTime - ((lUpTime \ 60) * 60))
    lUpTime = lUpTime \ 60
    upHoras = (lUpTime - ((lUpTime \ 24) * 24))
    lUpTime = lUpTime \ 24
    upDias = (lUpTime)
    
    
    'send_to_lcd control_char & "&h40&h" & Hex(32) & "&h30"
    
    send_to_lcd control_char & "&h42&h" & Hex(32) & "&h" & Hex(32)
    
    If Len(Trim(CStr(Ret))) = 3 Then
       send_to_lcd "Proc: " & CStr(Ret) + "%"
    Else
       send_to_lcd "Proc: " & SPACE(3 - Len(Trim(CStr(Ret)))) & CStr(Ret) + "%"
    End If

    
    send_to_lcd control_char & "&h42&h" & Hex(44) & "&h" & Hex(32)
    send_to_lcd Format(upDias, "000") & "-" & Format(upHoras, "00") & ":" & Format(upMinutos, "00") & ":" & Format(upSegundos, "00")

    lblUpTime.Caption = Format(upDias, "000") & "-" & Format(upHoras, "00") & ":" & Format(upMinutos, "00") & ":" & Format(upSegundos, "00")

    send_to_lcd control_char & "&h42&h" & Hex(32) & "&h" & Hex(33)
    If Len(Trim(CStr(Availp1))) = Len(Trim(CStr(Totp1))) Then
       send_to_lcd "Mem:  " & Availp1 & "/" & Totp1
       lblMemoria.Caption = Availp1 & "/" & Totp1
    Else
       send_to_lcd "Mem:  " & SPACE(Len(Trim(CStr(Totp1))) - Len(Trim(CStr(Availp1)))) & Availp1 & "/" & Totp1
       lblMemoria.Caption = SPACE(Len(Trim(CStr(Totp1))) - Len(Trim(CStr(Availp1)))) & Availp1 & "/" & Totp1
    End If
    
    
    send_to_lcd control_char & "&h42&h" & Hex(32) & "&h" & Hex(34)
    send_to_lcd "HD1:   " & Espaco & ds_Espaco
    lblHD1.Caption = Espaco & ds_Espaco
    
    send_to_lcd control_char & "&h42&h" & Hex(32) & "&h" & Hex(35)
    send_to_lcd "Temps: " & Dados_Sensores(0) & "° - " & Dados_Sensores(2) & "° - " & Dados_Sensores(4) & "°"
    lblTemperatura.Caption = Dados_Sensores(0) & "° - " & Dados_Sensores(2) & "° - " & Dados_Sensores(4) & "°"
    
    send_to_lcd control_char & "&h42&h" & Hex(32) & "&h" & Hex(36)
    send_to_lcd "Volts: " & Format(Dados_Sensores(34), "#0.0000") & "v   " & Format(Dados_Sensores(35), "#0.0000") & "v"
    lblVolts1.Caption = Format(Dados_Sensores(34), "#0.0000") & "v   " & Format(Dados_Sensores(35), "#0.0000") & "v"
    
    send_to_lcd control_char & "&h42&h" & Hex(32) & "&h" & Hex(37)
    send_to_lcd "      " & Format(Dados_Sensores(36), "#0.0000") & "v " & Format(Dados_Sensores(37), "#0.0000") & "v"
    lblVolts2.Caption = Format(Dados_Sensores(36), "#0.0000") & "v " & Format(Dados_Sensores(37), "#0.0000") & "v"
    
    send_to_lcd control_char & "&h42&h" & Hex(32) & "&h" & Hex(38)
    send_to_lcd "Cooler: " & Dados_Sensores(48) & " rpm"
    lblCooler.Caption = Dados_Sensores(48) & " rpm"
    
    
End Sub


