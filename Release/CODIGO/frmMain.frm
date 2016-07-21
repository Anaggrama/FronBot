VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   8685
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6750
      Top             =   1920
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer LagTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   8040
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   840
      Top             =   2880
   End
   Begin VB.Timer StaRecovery 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1800
      Top             =   2400
   End
   Begin VB.Timer Cuenta 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2880
   End
   Begin VB.Timer Barras 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   2400
   End
   Begin VB.Timer Test 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   2400
   End
   Begin VB.Timer Bot 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   2400
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   9000
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   4
      Top             =   3120
      Width           =   2280
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5760
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1800
      Visible         =   0   'False
      Width           =   8190
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   7080
      Top             =   2520
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   6600
      Top             =   2520
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5760
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6240
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4920
      Top             =   2520
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4200
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1725
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   3043
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":F172
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   8955
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3045
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblCambiarPj 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cambiar Personaje"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   8595
      Width           =   2295
   End
   Begin VB.Label lblDificultad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dificil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   10440
      TabIndex        =   16
      Top             =   8700
      Width           =   630
   End
   Begin VB.Label lblAimProm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   11490
      TabIndex        =   15
      Top             =   8250
      Width           =   255
   End
   Begin VB.Label lblFallados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   11520
      TabIndex        =   14
      Top             =   7830
      Width           =   105
   End
   Begin VB.Label lblAcertados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   11520
      TabIndex        =   13
      Top             =   7410
      Width           =   105
   End
   Begin VB.Label lblPromedio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9780
      TabIndex        =   12
      Top             =   8250
      Width           =   255
   End
   Begin VB.Label lblPerdidos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9780
      TabIndex        =   11
      Top             =   7830
      Width           =   105
   End
   Begin VB.Label lblGanados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   9780
      TabIndex        =   10
      Top             =   7410
      Width           =   105
   End
   Begin VB.Image imgMinimizar 
      Height          =   375
      Left            =   11280
      Top             =   15
      Width           =   375
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   11640
      Top             =   15
      Width           =   375
   End
   Begin VB.Image imgConfigBot 
      Height          =   255
      Left            =   3240
      Top             =   8550
      Width           =   1935
   End
   Begin VB.Image imgConfigTeclas 
      Height          =   255
      Left            =   6000
      Top             =   8550
      Width           =   1815
   End
   Begin VB.Label lblBotLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10065
      TabIndex        =   9
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblMana 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   9960
      TabIndex        =   2
      Top             =   6300
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10320
      MouseIcon       =   "frmMain.frx":F1F0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Image cmdInfo 
      Height          =   525
      Left            =   10560
      MouseIcon       =   "frmMain.frx":F342
      MousePointer    =   99  'Custom
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image cmdMoverHechiDown 
      Height          =   240
      Left            =   11430
      MouseIcon       =   "frmMain.frx":F494
      MousePointer    =   99  'Custom
      Top             =   3960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechiUp 
      Height          =   240
      Left            =   11430
      MouseIcon       =   "frmMain.frx":F5E6
      MousePointer    =   99  'Custom
      Top             =   3600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9990
      TabIndex        =   8
      Top             =   720
      Width           =   390
   End
   Begin VB.Image CmdLanzar 
      Height          =   495
      Left            =   8880
      MouseIcon       =   "frmMain.frx":F738
      MousePointer    =   99  'Custom
      Top             =   5400
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8880
      MouseIcon       =   "frmMain.frx":F88A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      Height          =   6240
      Left            =   270
      Top             =   2235
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Image InvEqu 
      Height          =   4230
      Left            =   8640
      Top             =   2235
      Width           =   2970
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   9840
      TabIndex        =   3
      Top             =   6810
      Width           =   1095
   End
   Begin VB.Shape shpMana 
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   9210
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Shape shpVida 
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   9210
      Top             =   6780
      Width           =   2295
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public MoviendoVida As Long
Public MoviendoMana As Long
Public ActualHP As Integer
Public ActualMP As Integer

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Private BotonConfigTeclas As clsGraphicalButton
Private BotonConfigBot As clsGraphicalButton
Private BotonSalir As clsGraphicalButton
Private BotonMinimizar As clsGraphicalButton
Private BotonMoveHechiUp As clsGraphicalButton
Private BotonMoveHechiDown As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public picSkillStar As Picture

Dim PuedeMacrear As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Barras_Timer
' Author    : Anagrama
' Date      : ???
' Purpose   : Maneja las barras de vida y mana dinámicas, según cuanto mayor sea la perdida o ganancia, mas rápido se mueven.
'---------------------------------------------------------------------------------------
'
Private Sub Barras_Timer()
    Dim bWidth As Byte

    If MoviendoVida > 0 Then 'Si hay cambias en la vida.
        If charlist(UserCharIndex).MinHP < ActualHP Then 'Pierde.
            ActualHP = ActualHP - MoviendoVida
            If ActualHP < charlist(UserCharIndex).MinHP Then ActualHP = charlist(UserCharIndex).MinHP
            bWidth = (((ActualHP / 100) / (charlist(UserCharIndex).MaxHP / 100)) * 153)
            frmMain.shpVida.Width = 153 - bWidth
            frmMain.shpVida.Left = 614 + (153 - frmMain.shpVida.Width)
            frmMain.shpVida.Visible = (bWidth <> 153)
        ElseIf charlist(UserCharIndex).MinHP > ActualHP Then 'Gana.
            ActualHP = ActualHP + MoviendoVida
            If ActualHP > charlist(UserCharIndex).MinHP Then ActualHP = charlist(UserCharIndex).MinHP
            bWidth = (((ActualHP / 100) / (charlist(UserCharIndex).MaxHP / 100)) * 153)
            frmMain.shpVida.Width = 153 - bWidth
            frmMain.shpVida.Left = 614 + (153 - frmMain.shpVida.Width)
            frmMain.shpVida.Visible = (bWidth <> 153)
        Else: MoviendoVida = 0 'Termino de mover.
        End If
    End If
    If MoviendoMana > 0 Then 'Si hay cambios en la mana.
        If charlist(UserCharIndex).MinMAN < ActualMP Then
            ActualMP = ActualMP - MoviendoMana
            If ActualMP < charlist(UserCharIndex).MinMAN Then ActualMP = charlist(UserCharIndex).MinMAN
            bWidth = (((ActualMP / 100) / (charlist(UserCharIndex).MaxMAN / 100)) * 153)
            frmMain.shpMana.Width = 153 - bWidth
            frmMain.shpMana.Left = 614 + (153 - frmMain.shpMana.Width)
            frmMain.shpMana.Visible = (bWidth <> 153)
        ElseIf charlist(UserCharIndex).MinMAN > ActualMP Then
            ActualMP = ActualMP + MoviendoMana
            If ActualMP > charlist(UserCharIndex).MinMAN Then ActualMP = charlist(UserCharIndex).MinMAN
            bWidth = (((ActualMP / 100) / (charlist(UserCharIndex).MaxMAN / 100)) * 153)
            frmMain.shpMana.Width = 153 - bWidth
            frmMain.shpMana.Left = 614 + (153 - frmMain.shpMana.Width)
            frmMain.shpMana.Visible = (bWidth <> 153)
        Else: MoviendoMana = 0
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Bot_Timer
' Author    : Anagrama
' Date      : ???
' Purpose   : Checkea cada contador de intervalos de los bots y revisa que acción tomar segun cada situación.
' Actualizado para que cada ia tenga su propio sub asi es mas facil de leer y modificar.
'---------------------------------------------------------------------------------------
'
Private Sub Bot_Timer()
    Dim i As Integer
    
    If EnCuenta Then Exit Sub

    For i = 1 To MaxUsers
        If charlist(i).Bot = 1 Then
            Select Case charlist(i).Ai
                Case eBotAi.Mago
                    Call MagoAI(i)
                Case eBotAi.Guerrero
                    Call GuerreroAI(i)
                Case eBotAi.Clerigo, eBotAi.Bardo 'Usan la misma ia
                    Call ClerigoAI(i)
                Case eBotAi.Paladin, eBotAi.Asesino 'Usan la misma ia
                    Call PaladinAI(i)
                Case eBotAi.Druida
                    Call DruidaAI(i)
            End Select
        Else
            If charlist(i).Quieto < 50 Then charlist(i).Quieto = charlist(i).Quieto + 1
        End If
    Next i

End Sub

Private Sub cmdCambiarPj_Click()
    frmPersonaje.Modo = 2
    frmPersonaje.Show
End Sub

Private Sub cmdMoverHechiDown_Click()
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
        If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

        sTemp = hlst.List(hlst.ListIndex + 1)
        hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
        hlst.List(hlst.ListIndex) = sTemp
        hlst.ListIndex = hlst.ListIndex + 1
    End If
End Sub

Private Sub cmdMoverHechiUp_Click()
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
        If hlst.ListIndex = 0 Then Exit Sub

        sTemp = hlst.List(hlst.ListIndex - 1)
        hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
        hlst.List(hlst.ListIndex) = sTemp
        hlst.ListIndex = hlst.ListIndex - 1
    End If
End Sub

Private Sub Form_Load()
    
    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If

    Me.Picture = LoadPicture(DirGraficos & "Principal Bot.JPG")
    
    InvEqu.Picture = LoadPicture(DirGraficos & "InventarioBot.jpg")
    
    Call LoadButtons
    
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub LoadButtons()

    Set BotonConfigTeclas = New clsGraphicalButton
    Set BotonConfigBot = New clsGraphicalButton
    Set BotonSalir = New clsGraphicalButton
    Set BotonMinimizar = New clsGraphicalButton
    Set BotonMoveHechiUp = New clsGraphicalButton
    Set BotonMoveHechiDown = New clsGraphicalButton
    
    Call BotonConfigTeclas.Initialize(imgConfigTeclas, App.path & "\Graficos\configurarteclas_normal.jpg", App.path & "\Graficos\configurarteclas_hover.jpg", App.path & "\Graficos\configurarteclas_click.jpg", Me)
    Call BotonConfigBot.Initialize(imgConfigBot, App.path & "\Graficos\configuracion_normal.jpg", App.path & "\Graficos\configuracion_hover.jpg", App.path & "\Graficos\configuracion_click.jpg", Me)
    Call BotonSalir.Initialize(imgSalir, App.path & "\Graficos\cruzcerrar.jpg", App.path & "\Graficos\cruzcerrarhover.jpg", App.path & "\Graficos\cruzcerrarclick.jpg", Me)
    Call BotonMinimizar.Initialize(imgMinimizar, App.path & "\Graficos\menosminimizar.jpg", App.path & "\Graficos\menosminimizarhover.jpg", App.path & "\Graficos\menosminimizarclick.jpg", Me)
    Call BotonMoveHechiUp.Initialize(cmdMoverHechiUp, App.path & "\Graficos\flechahechizos2.jpg", App.path & "\Graficos\flechahechizos2hover.jpg", App.path & "\Graficos\flechahechizos2click.jpg", Me)
    Call BotonMoveHechiDown.Initialize(cmdMoverHechiDown, App.path & "\Graficos\flechahechizos.jpg", App.path & "\Graficos\flechahechizoshover.jpg", App.path & "\Graficos\flechahechizosclick.jpg", Me)
    
    Set LastPressed = New clsGraphicalButton

    imgSalir.MouseIcon = picMouseIcon
    imgMinimizar.MouseIcon = picMouseIcon
End Sub

Private Sub cmdMoverHechi_Click(index As Integer)
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case index
            Case 1 'subir
                If hlst.ListIndex = 0 Then Exit Sub
            Case 0 'bajar
                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
        
        Select Case index
            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1
            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
'***************************************************
#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
    
    If (Not SendTxt.Visible) Then

        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    Call UsarItem(0)
            End Select
        Else
            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyRPU)
            Call WriteRequestPositionUpdate
            
        Case CustomKeys.BindedKey(eKeyType.mKeyEquip)
            If charlist(UserCharIndex).MinHP > 0 Then
                Call EquiparItem(UserCharIndex, Inventario.SelectedItem)
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Estás muerto!!.", 65, 190, 156)
            End If
                    
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If charlist(UserCharIndex).MinHP = 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Estás muerto!!.", 65, 190, 156)
                Exit Sub
            End If
            Call WriteAttack
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendTxt.Visible = False Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            Else
                If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                If picInv.Visible Then
                    picInv.SetFocus
                Else
                    hlst.SetFocus
                End If
            End If
            
        Case CustomKeys.BindedKey(eKeyType.mKeyPausa)
            If SendTxt.Visible = True Then Exit Sub
            
            If Hosting = 1 Then
                If pausa = True Then
                    pausa = False
                    Bot.Enabled = True
                    Call ServerSendData(SendTarget.ToAll, UserCharIndex, PrepareMessagePauseToggle(0))
                Else
                    pausa = True
                    Bot.Enabled = False
                    Call ServerSendData(SendTarget.ToAll, UserCharIndex, PrepareMessagePauseToggle(1))
                End If
            End If
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Cuenta_Timer()
    If CuentaR > 0 Then
        Call ServerSendData(SendTarget.ToAll, UserCharIndex, PrepareMessageConsoleMsg("El combate iniciará en " & CuentaR & " segundos.", FontTypeNames.FONTTYPE_PARTY))
        CuentaR = CuentaR - 1
    Else
        Call ServerSendData(SendTarget.ToAll, UserCharIndex, PrepareMessageConsoleMsg("¡El combate ha iniciado!.", FontTypeNames.FONTTYPE_PARTY))
        Cuenta.Enabled = False
        Call ServerSendData(SendTarget.ToAll, UserCharIndex, PrepareMessageCuentaToggle(0))
    End If
End Sub

Private Sub imgConfigBot_Click()
    If Hosting = 1 Then
        frmStart.Show
    End If
End Sub

Private Sub imgConfigTeclas_Click()
    frmCustomKeys.Show
End Sub

Private Sub imgMinimizar_Click()
    Me.WindowState = 1
End Sub

Private Sub imgSalir_Click()
    Call LimpiaWsApi
    If frmMain.Socket1.Connected Then
        Call WriteQuit
        Call FlushBuffer
    End If
    End
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub LagTimer_Timer()
    Dim i As Byte
    Dim a As Byte
    If (Not Pendiente) <> -1 Then
        For i = 1 To UBound(Pendiente)
            If i > UBound(Pendiente) Then Exit Sub
            If Pendiente(i).Slot > 0 Then
                Call charlist(Pendiente(i).Slot).incomingData.WriteBlock(Pendiente(i).Datos)
                Call ServerHandleIncomingData(Pendiente(i).Slot)
                Pendiente(i).Slot = 0
                If i < UBound(Pendiente) Then
                    For a = i To UBound(Pendiente) - 1
                        Pendiente(a).Slot = Pendiente(a + 1).Slot
                        Pendiente(a).Datos = Pendiente(a + 1).Datos
                    Next a
                    ReDim Preserve Pendiente(1 To UBound(Pendiente) - 1) As tPendiente
                End If
            End If
        Next i
    End If
End Sub

Private Sub lblCambiarPj_Click()
    frmPersonaje.Modo = 2
    frmPersonaje.Show
End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo ErrHandler:
    Dim i As Long
    
    For i = 1 To MaxUsers
        If charlist(i).ConnIDValida Then
            If charlist(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, charlist(i).outgoingData.ReadASCIIStringFixed(charlist(i).outgoingData.length))
            End If
        End If
    Next i

Exit Sub

ErrHandler:
    Resume Next
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    'If KeyCode = vbKeyReturn Then
    '    If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
    '
    '    stxtbuffer = ""
    '    SendTxt.Text = ""
    '    KeyCode = 0
    '    SendTxt.Visible = False
    '
    '    If picInv.Visible Then
    '        picInv.SetFocus
    '    Else
    '        hlst.SetFocus
    '    End If
    'End If
End Sub

Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
    Second.Enabled = True

    Call Login

End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long

    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).name <> frmLogin.name And Forms(i).name <> Me.name Then
            Unload Forms(i)
        End If
    Loop
    
    On Local Error GoTo 0
    
    frmLogin.MousePointer = vbNormal
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmLogin.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If Not frmLogin.Visible Then
        frmLogin.Show
    End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub

    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call ClientHandleIncomingData
End Sub


Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub UsarItem(ByVal Click As Byte)
    Dim bWidth As Long
    
    If pausa Or EnCuenta Then Exit Sub
    If charlist(UserCharIndex).MinHP = 0 Then
        Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Estás muerto!!.", 65, 190, 156)
        Exit Sub
    End If
    If (Inventario.SelectedItem = 1) Then
        Call WriteUseItem(Inventario.OBJIndex(Inventario.SelectedItem), Click)
    ElseIf (Inventario.SelectedItem = 5) Or (Inventario.SelectedItem = 2) Then
        If Inventario.OBJIndex(Inventario.SelectedItem) = 4 Then
            If charlist(UserCharIndex).TipoArma = 2 Then
                Me.MousePointer = 2
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, "Debes equipar el arco antes de usarlo.", 65, 190, 156)
            End If
        Else
            Call WriteUseItem(Inventario.OBJIndex(Inventario.SelectedItem), Click)
        End If
    End If

End Sub

Private Sub cmdLanzar_Click()
    If hlst.ListIndex < 1 Then Exit Sub
    If hlst.List(hlst.ListIndex) <> "(Nada)" Then
        MySpell = hlst.List(hlst.ListIndex)
        Me.MousePointer = 2
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub DespInv_Click(index As Integer)
    Inventario.ScrollInventory (index = 0)
End Sub

Private Sub Form_Click()
    Dim hI As Byte
    
    If pausa Or EnCuenta Then Exit Sub
    
    If Not InGameArea Then Exit Sub
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    Call WriteLeftClick(tX, tY)
    
    If Me.MousePointer = 2 Then
        If charlist(UserCharIndex).Ai <> eBotAi.Guerrero Then
            Select Case MySpell
                Case "Apocalípsis"
                    hI = 1
                Case "Descarga Eléctrica"
                    hI = 2
                Case "Inmovilizar"
                    hI = 3
                Case "Devolver Movilidad"
                    hI = 4
                Case "Tormenta de Fuego"
                    hI = 5
                Case "Resucitar"
                    hI = 6
            End Select
            If hI > 0 And hI < 7 Then Call WriteCastSpell(tX, tY, hI)
            Me.MousePointer = 1
        Else
            If charlist(UserCharIndex).MinHP = 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Estás muerto!!.", 65, 190, 156)
                Exit Sub
            End If
            
            Call WriteLanzaFlecha(tX, tY)
            
            Me.MousePointer = 1
        End If
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewShp.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
    
    LastPressed.ToggleToNormal
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.path & "\Graficos\inventariobot.jpg")

    ' Activo controles de inventario
    picInv.Visible = True

    ' Desactivo controles de hechizo
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechiDown.Visible = False
    cmdMoverHechiUp.Visible = False
    
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.path & "\Graficos\magiabot.jpg")
    
    ' Activo controles de hechizos
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechiDown.Visible = True
    cmdMoverHechiUp.Visible = True
    
    ' Desactivo controles de inventario
    picInv.Visible = False

End Sub

Private Sub picInv_DblClick()
    Call UsarItem(1)
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then





#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim M As New frmMenuseFashion
            
            Load M
            M.SetCallback Me
            M.SetMenuId 1
            M.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                M.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                M.ListaSetItem 0, "<NPC>", True
            End If
            M.ListaSetItem 1, "Comerciar"
            
            M.ListaFin
            M.Show , Me

        End If
    End If
End If

#End If
End Sub


Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + MainViewShp.Width Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + MainViewShp.Height Then Exit Function
    
    InGameArea = True
End Function

Private Sub StaRecovery_Timer()
    Dim i As Integer
    Dim massta As Integer
    For i = 1 To MaxUsers
        With charlist(i)
            If .ConnID <> -1 Or .Bot = 1 Then
                If .MinHP > 0 Then
                    If .MinSTA < .MaxSTA Then
                        massta = RandomNum(.MaxSTA * 5 / 100)
                        .MinSTA = .MinSTA + massta
                        If .MinSTA > .MaxSTA Then
                            .MinSTA = .MaxSTA
                        End If
                        If .Bot = 0 Then
                            Call WriteUpdateCharStats(i)
                        End If
                    End If
                End If
            End If
        End With
    Next
End Sub

Private Sub Test_Timer()

    If CharTimer(1).Check(TimersIndex.GolpeU, False) Then
        If CharTimer(1).Check(TimersIndex.UseItem, False) Then
            If charlist(2).Lanzando = 0 And CharTimer(1).Check(TimersIndex.WaitP, False) Then
                If CharTimer(1).Check(TimersIndex.UseItemWithDblClick, False) Then
                    Call CharTimer(1).Restart(TimersIndex.UseItem)
                    Call CharTimer(1).Restart(TimersIndex.UseItemWithDblClick)
                    Call BotTomaPot(2)
                End If
            Else
                If CharTimer(1).Check(TimersIndex.UseItemWithU, False) Then
                    Call CharTimer(1).Restart(TimersIndex.UseItem)
                    Call CharTimer(1).Restart(TimersIndex.UseItemWithU)
                    Call BotTomaPot(2)
                End If
            End If
        End If
    End If
    
End Sub

Private Sub Winsock2_Connect()
#If SeguridadAlkon = 1 Then
    Call modURL.ProcessRequest
#End If
End Sub
