VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FronBot"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSound 
      BackColor       =   &H00000000&
      Caption         =   "Sonido/FX"
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   330
      TabIndex        =   9
      Top             =   3750
      Width           =   1305
   End
   Begin FronBot.lvButtons_H cmdCrear 
      Default         =   -1  'True
      Height          =   675
      Left            =   300
      TabIndex        =   0
      Top             =   2340
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   1191
      Caption         =   "&Nueva Partida"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3930
      TabIndex        =   3
      Text            =   "Jugador"
      Top             =   3570
      Width           =   1455
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3930
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   3210
      Width           =   1455
   End
   Begin FronBot.lvButtons_H cmdConectar 
      Height          =   705
      Left            =   2490
      TabIndex        =   1
      Top             =   2340
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   1244
      Caption         =   "&Unirse a una partida"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Este es un cliente/servidor de agite, que permite combinar bots inteligentes con usuarios reales mediante internet o lan."
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
      Height          =   735
      Left            =   300
      TabIndex        =   8
      Top             =   1020
      Width           =   5010
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Bot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   435
      Left            =   1080
      TabIndex        =   7
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Fron"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   300
      TabIndex        =   6
      Top             =   240
      Width           =   825
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Nombre:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3090
      TabIndex        =   5
      Top             =   3570
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "IP:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3090
      TabIndex        =   4
      Top             =   3210
      Width           =   255
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSound_Click()
    If chkSound.value = 1 Then
        Audio.SoundActivated = True
    Else
        Audio.SoundActivated = False
    End If
End Sub

Private Sub cmdConectar_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmPersonaje.Modo = 0
    frmPersonaje.Show
End Sub

Private Sub cmdCrear_Click()
    Call Audio.PlayWave(SND_CLICK)
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    Hosting = 1
    frmStart.Show
    Me.Visible = False
End Sub

Private Sub Form_Load()
    If Audio.SoundActivated Then
        chkSound.value = 1
    Else
        chkSound.value = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
