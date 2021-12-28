VERSION 5.00
Begin VB.Form frmPersonaje 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FronBot"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3405
   Icon            =   "frmPersonaje.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3405
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSexo 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   930
      Width           =   2000
   End
   Begin VB.ComboBox cmbRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   570
      Width           =   2000
   End
   Begin VB.ComboBox cmbClase 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      ItemData        =   "frmPersonaje.frx":F172
      Left            =   1260
      List            =   "frmPersonaje.frx":F174
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   2000
   End
   Begin VB.ComboBox cmbUserLvl 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1290
      Width           =   2000
   End
   Begin FronBot.lvButtons_H cmbAceptar 
      Height          =   675
      Left            =   210
      TabIndex        =   8
      Top             =   2850
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1191
      Caption         =   "&Confirmar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8454016
      cFHover         =   8438015
      cBhover         =   0
      LockHover       =   2
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   16384
   End
   Begin FronBot.lvButtons_H cmdBlue 
      Height          =   675
      Left            =   240
      TabIndex        =   9
      Top             =   1890
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1191
      Caption         =   "Equipo Azul"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16744576
      cFHover         =   8438015
      cBhover         =   0
      LockHover       =   2
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   8388608
   End
   Begin FronBot.lvButtons_H cmdRed 
      Height          =   675
      Left            =   1470
      TabIndex        =   10
      Top             =   1890
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1191
      Caption         =   "Equipo Rojo"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421631
      cFHover         =   8438015
      cBhover         =   0
      LockHover       =   2
      cGradient       =   0
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   64
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1350
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   990
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   270
      Width           =   570
   End
End
Attribute VB_Name = "frmPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Modo As Byte
Public Team As Byte ' 1 blue 2 red

Private Sub cmbAceptar_Click()
    Select Case Modo
        Case 0 'Login cliente
            UserClase = cmbClase.ListIndex
            UserRaza = cmbRaza.ListIndex
            UserGenero = cmbSexo.ListIndex
            UserLvl = cmbUserLvl.ListIndex + 35
            UserName = frmLogin.txtNombre.Text
            UserTeam = Team
            
            If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
                DoEvents
            End If
            Hosting = 0
            
            frmMain.Socket1.HostName = frmLogin.txtIP.Text
            frmMain.Socket1.RemotePort = 7600
            frmMain.Socket1.connect
        Case 1 'Login host
            UserClase = cmbClase.ListIndex
            UserRaza = cmbRaza.ListIndex
            UserGenero = cmbSexo.ListIndex
            UserLvl = cmbUserLvl.ListIndex + 35
            UserTeam = 1
            UserName = frmLogin.txtNombre.Text
            
            Call IniBot(frmStart.cmbBalance.List(frmStart.cmbBalance.ListIndex))
            
            Call SecurityIp.InitIpTables(1000)
            Call IniciaWsApi(frmMain.hwnd)
            SockListen = ListenForConnect(7600, hWndMsg, "")
            
            If SockListen = INVALID_SOCKET Then
                MsgBox "Cierre y vuelva a abrir el cliente."
                Exit Sub
            End If
            
            frmMain.Socket1.HostName = "127.0.0.1"
            frmMain.Socket1.RemotePort = 7600
            frmMain.Socket1.connect
        Case 2 'Cambio de personaje
            UserClase = cmbClase.ListIndex
            UserRaza = cmbRaza.ListIndex
            UserGenero = cmbSexo.ListIndex
            UserLvl = cmbUserLvl.ListIndex + 35
            UserTeam = Team
            
            Call WriteChangePj
            Me.Visible = False
        Case 3 'Agregar bot
            UserTeam = Team
            If UserTeam = 1 Then 'Equipo 1 o azul
                If frmStart.lstTeam1.ListCount >= 5 Then
                    Exit Sub
                ElseIf UserCharIndex = 0 And frmStart.lstTeam1.ListCount >= 4 Then
                    Exit Sub
                End If
                TeamData1(frmStart.lstTeam1.ListCount + 1).Clase = cmbClase.ListIndex
                TeamData1(frmStart.lstTeam1.ListCount + 1).Raza = cmbRaza.ListIndex
                TeamData1(frmStart.lstTeam1.ListCount + 1).Genero = cmbSexo.ListIndex
                TeamData1(frmStart.lstTeam1.ListCount + 1).Nivel = cmbUserLvl.ListIndex + 35
                TeamData1(frmStart.lstTeam1.ListCount + 1).Bot = 1
                frmStart.lstTeam1.AddItem "Bot " & frmStart.lstTeam1.ListCount + 1
            Else 'Equipo 2 o rojo
                If frmStart.lstTeam2.ListCount >= 5 Then Exit Sub
                TeamData2(frmStart.lstTeam2.ListCount + 1).Clase = cmbClase.ListIndex
                TeamData2(frmStart.lstTeam2.ListCount + 1).Raza = cmbRaza.ListIndex
                TeamData2(frmStart.lstTeam2.ListCount + 1).Genero = cmbSexo.ListIndex
                TeamData2(frmStart.lstTeam2.ListCount + 1).Nivel = cmbUserLvl.ListIndex + 35
                TeamData2(frmStart.lstTeam2.ListCount + 1).Bot = 1
                frmStart.lstTeam2.AddItem "Bot " & frmStart.lstTeam2.ListCount + 1
            End If
            Unload Me
    End Select
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub cmdBlue_Click()
    Team = 1
    Call TeamChoice
End Sub

Private Sub cmdRed_Click()
    If Modo <> 1 Then
        Team = 2
        Call TeamChoice
    End If
End Sub

Private Sub Form_Load()
    Team = 1
    Call TeamChoice
    If Modo = 1 Then
        cmdRed.Enabled = False
     Else
        cmdRed.Enabled = True
    End If
    cmbClase.AddItem "Mago"
    cmbClase.AddItem "Paladín"
    cmbClase.AddItem "Clérigo"
    cmbClase.AddItem "Asesino"
    cmbClase.AddItem "Bardo"
    cmbClase.AddItem "Druida"
    cmbClase.AddItem "Guerrero"
    cmbClase.ListIndex = 0
    cmbRaza.AddItem "Humano"
    cmbRaza.AddItem "Gnomo"
    cmbRaza.AddItem "Elfo"
    cmbRaza.AddItem "Enano"
    cmbRaza.AddItem "Elfo Drow"
    cmbRaza.ListIndex = 0
    cmbSexo.AddItem "Hombre"
    cmbSexo.AddItem "Mujer"
    cmbSexo.ListIndex = 0
    cmbUserLvl.AddItem "35"
    cmbUserLvl.AddItem "36"
    cmbUserLvl.AddItem "37"
    cmbUserLvl.AddItem "38"
    cmbUserLvl.AddItem "39"
    cmbUserLvl.AddItem "40"
    cmbUserLvl.AddItem "41"
    cmbUserLvl.AddItem "42"
    cmbUserLvl.AddItem "43"
    cmbUserLvl.AddItem "44"
    cmbUserLvl.AddItem "45"
    cmbUserLvl.AddItem "46"
    cmbUserLvl.AddItem "47"
    cmbUserLvl.AddItem "48"
    cmbUserLvl.AddItem "49"
    cmbUserLvl.AddItem "50"
    cmbUserLvl.ListIndex = 5
    
    Select Case Modo
        Case 0, 1, 2
            Me.Caption = "Seleccionar personaje"
            cmbAceptar.Caption = "Seleccionar"
        Case 3
            Me.Caption = "Agregar bot"
            cmbAceptar.Caption = "Agregar bot"
    End Select
End Sub

Private Sub TeamChoice()
    If UserTeam <> 1 Then
        cmdBlue.BackColor = &H800000
        cmdBlue.ForeColor = &HFF8080
        cmdRed.BackColor = &HC0&
        cmdRed.ForeColor = vbWhite
    Else
        cmdBlue.BackColor = &HFF0000
        cmdBlue.ForeColor = vbWhite
        cmdRed.BackColor = &H40&
        cmdRed.ForeColor = &H8080FF
    End If
End Sub
