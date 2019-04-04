VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FronBot"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "Jugador"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdConectar 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear partida"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "IP:"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConectar_Click()
    frmPersonaje.Modo = 0
    frmPersonaje.Show
End Sub

Private Sub cmdCrear_Click()
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    Hosting = 1
    frmStart.Show
    Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
