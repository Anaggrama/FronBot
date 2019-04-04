VERSION 5.00
Begin VB.Form frmPersonaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FronBot"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2325
   Icon            =   "frmPersonaje.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cmbTeam 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cmbSexo 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox cmbRaza 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox cmbClase 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cmbUserLvl 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Equipo:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nivel:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Genero:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Raza:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clase:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Modo As Byte

Private Sub cmbAceptar_Click()
    Select Case Modo
        Case 0 'Login cliente
            UserClase = cmbClase.ListIndex
            UserRaza = cmbRaza.ListIndex
            UserGenero = cmbSexo.ListIndex
            UserLvl = cmbUserLvl.ListIndex + 35
            UserTeam = cmbTeam.ListIndex + 1
            UserName = frmLogin.txtNombre.Text
            
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
            Call IniciaWsApi(frmMain.hWnd)
            SockListen = ListenForConnect(7600, hWndMsg, "")
            
            frmMain.Socket1.HostName = "127.0.0.1"
            frmMain.Socket1.RemotePort = 7600
            frmMain.Socket1.connect
        Case 2 'Cambio de personaje
            UserClase = cmbClase.ListIndex
            UserRaza = cmbRaza.ListIndex
            UserGenero = cmbSexo.ListIndex
            UserLvl = cmbUserLvl.ListIndex + 35
            UserTeam = cmbTeam.ListIndex + 1
            
            Call WriteChangePj
            Me.Visible = False
        Case 3 'Agregar bot
            If cmbTeam.ListIndex = 0 Then 'Equipo 1 o azul
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
    End Select
End Sub

Private Sub Form_Load()
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
    cmbTeam.AddItem "Azul"
    cmbTeam.AddItem "Rojo"
    cmbTeam.ListIndex = 0
    
    Select Case Modo
        Case 0, 1, 2
            Me.Caption = "Seleccionar personaje"
        Case 3
            Me.Caption = "Agregar bot"
    End Select
End Sub
