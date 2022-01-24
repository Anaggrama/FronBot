VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FronBot"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRandomAi 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   5160
      Width           =   255
   End
   Begin VB.ListBox lstTeam2 
      BackColor       =   &H00000000&
      ForeColor       =   &H008080FF&
      Height          =   2400
      Left            =   2550
      TabIndex        =   10
      Top             =   1290
      Width           =   2265
   End
   Begin VB.ListBox lstTeam1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF8080&
      Height          =   2400
      Left            =   150
      TabIndex        =   9
      Top             =   1290
      Width           =   2265
   End
   Begin VB.CheckBox chkResuVida 
      BackColor       =   &H00000000&
      Caption         =   "Check1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   6180
      Width           =   255
   End
   Begin VB.CheckBox chkResu 
      BackColor       =   &H00000000&
      Caption         =   "Check1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   5730
      Width           =   255
   End
   Begin VB.ComboBox cmbBalance 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4110
      Width           =   1575
   End
   Begin VB.CheckBox chkRandom 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1950
      TabIndex        =   1
      Top             =   4650
      Width           =   255
   End
   Begin VB.ComboBox cmbDificultad 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      ItemData        =   "frmStart.frx":F172
      Left            =   1185
      List            =   "frmStart.frx":F174
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4110
      Width           =   1335
   End
   Begin FronBot.lvButtons_H cmdAgregar 
      Height          =   675
      Left            =   4920
      TabIndex        =   20
      Top             =   2250
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1191
      Caption         =   "&Nuevo Bot"
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
   Begin FronBot.lvButtons_H cmdRemover 
      Height          =   675
      Left            =   4920
      TabIndex        =   21
      Top             =   3030
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1191
      Caption         =   "&Quitar Bot"
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
      cBack           =   64
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Inteligencia Aleatoria:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   23
      Top             =   5190
      Width           =   2265
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo Rojo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   6
      Left            =   2550
      TabIndex        =   19
      Top             =   930
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo Azul:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   18
      Top             =   930
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Configuración de nueva partida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   210
      TabIndex        =   17
      Top             =   270
      Width           =   5445
   End
   Begin VB.Label lblTeam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1980
      TabIndex        =   16
      Top             =   1290
      Width           =   45
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dificultad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   15
      Top             =   4110
      Width           =   975
   End
   Begin VB.Label lblNivel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1950
      TabIndex        =   14
      Top             =   2850
      Width           =   45
   End
   Begin VB.Label lblGenero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1950
      TabIndex        =   13
      Top             =   2610
      Width           =   45
   End
   Begin VB.Label lblRaza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1950
      TabIndex        =   12
      Top             =   2370
      Width           =   45
   End
   Begin VB.Label lblClase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1980
      TabIndex        =   11
      Top             =   1530
      Width           =   45
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Resu no saca vida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   6180
      Width           =   1665
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Resu:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   5730
      Width           =   945
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Aspecto Aleatorio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   4
      Top             =   4680
      Width           =   2085
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   4110
      Width           =   855
   End
   Begin VB.Image imgCruz 
      Height          =   375
      Left            =   8160
      Top             =   150
      Width           =   375
   End
   Begin VB.Image imgPelear 
      Height          =   705
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2235
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmStart
' Author    : Anagrama
' Date      : ???
' Purpose   : Formulario inicial del programa para personalizar los bots, el usuario e iniciar el combate.
'---------------------------------------------------------------------------------------

Option Explicit

Private BotonPelear As clsGraphicalButton
Private BotonCruz As clsGraphicalButton
Public LastPressed As clsGraphicalButton

Private ListSelected As Byte

Private Sub cmdAgregar_Click()
    Call Audio.PlayWave(SND_CLICK)
    frmPersonaje.Modo = 3
    frmPersonaje.Show
End Sub

Private Sub cmdRemover_Click()
    If ListSelected = 0 Then
        MsgBox "Debes seleccionar un bot primero."
        Exit Sub
    ElseIf ListSelected = 1 Then
        If lstTeam1.ListIndex >= 0 Then
            If TeamData1(lstTeam1.ListIndex + 1).index > 0 Then
                If TeamData1(lstTeam1.ListIndex + 1).Bot = 0 Then
                    MsgBox "No estás seleccionando un bot."
                    Exit Sub
                End If
            End If
            If lstTeam1.ListIndex + 1 < 5 Then _
                Call PushTeamData(1, lstTeam1.ListIndex + 1)
            lstTeam1.RemoveItem lstTeam1.ListIndex
        End If
    Else
        If lstTeam2.ListIndex >= 0 Then
            If TeamData2(lstTeam2.ListIndex + 1).index > 0 Then
                If TeamData2(lstTeam2.ListIndex + 1).Bot = 0 Then
                    MsgBox "No estás seleccionando un bot."
                    Exit Sub
                End If
            End If
            If lstTeam2.ListIndex + 1 < 5 Then _
                Call PushTeamData(2, lstTeam2.ListIndex + 1)
            lstTeam2.RemoveItem lstTeam2.ListIndex
        End If
    End If
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
    Dim BalanceFiles As String
    Dim i As Byte
    
    Set BotonPelear = New clsGraphicalButton
    Set BotonCruz = New clsGraphicalButton
    Set LastPressed = New clsGraphicalButton
    
    Call BotonPelear.Initialize(imgPelear, App.path & "\Graficos\Boton_Combatir_Normal.jpg", _
                                    App.path & "\Graficos\Boton_Combatir_Hover.jpg", _
                                    App.path & "\Graficos\Boton_Combatir_Click.jpg", Me)
                                    
    Call BotonCruz.Initialize(imgCruz, App.path & "\Graficos\CruzCerrar.jpg", _
                                    App.path & "\Graficos\CruzCerrarHover.jpg", _
                                    App.path & "\Graficos\CruzCerrarClick.jpg", Me)
                                    
    lstTeam1.Clear
    lstTeam2.Clear
    If UBound(Team1) > 0 Then
        If Team1(1) = 0 Then
            For i = 1 To 5
                TeamData1(i).Clase = eBotAi.Mago
                TeamData1(i).Raza = eRazaAi.Humano
                TeamData1(i).Genero = 0
                TeamData1(i).Nivel = 40
                TeamData1(i).Bot = 0
                TeamData1(i).index = 0
            Next i
        Else
            For i = 1 To UBound(Team1)
                TeamData1(i).Clase = charlist(Team1(i)).Ai
                TeamData1(i).Raza = charlist(Team1(i)).Raza
                TeamData1(i).Genero = charlist(Team1(i)).Genero
                TeamData1(i).Nivel = charlist(Team1(i)).Lvl
                TeamData1(i).Bot = charlist(Team1(i)).Bot
                TeamData1(i).index = Team1(i)
                lstTeam1.AddItem charlist(Team1(i)).Nombre
            Next i
            If UBound(Team1) < 5 Then
                For i = UBound(Team1) + 1 To 5
                    TeamData1(i).Clase = eBotAi.Mago
                    TeamData1(i).Raza = eRazaAi.Humano
                    TeamData1(i).Genero = 0
                    TeamData1(i).Nivel = 40
                    TeamData1(i).Bot = 0
                    TeamData1(i).index = 0
                Next i
            End If
        End If
    End If
    If UBound(Team2) > 0 Then
        If Team2(1) = 0 Then
            For i = 1 To 5
                TeamData2(i).Clase = eBotAi.Mago
                TeamData2(i).Raza = eRazaAi.Humano
                TeamData2(i).Genero = 0
                TeamData2(i).Nivel = 40
                TeamData2(i).Bot = 0
                TeamData2(i).index = 0
            Next i
        Else
            For i = 1 To UBound(Team2)
                TeamData2(i).Clase = charlist(Team2(i)).Ai
                TeamData2(i).Raza = charlist(Team2(i)).Raza
                TeamData2(i).Genero = charlist(Team2(i)).Genero
                TeamData2(i).Nivel = charlist(Team2(i)).Lvl
                TeamData2(i).Bot = charlist(Team2(i)).Bot
                TeamData2(i).index = Team2(i)
                lstTeam2.AddItem charlist(Team2(i)).Nombre
            Next i
            If UBound(Team2) < 5 Then
                For i = UBound(Team2) + 1 To 5
                    TeamData2(i).Clase = eBotAi.Mago
                    TeamData2(i).Raza = eRazaAi.Humano
                    TeamData2(i).Genero = 0
                    TeamData2(i).Nivel = 40
                    TeamData2(i).Bot = 0
                    TeamData2(i).index = 0
                Next i
            End If
        End If
    End If
    cmbDificultad.AddItem "Imposible"
    cmbDificultad.AddItem "Muy Dificil"
    cmbDificultad.AddItem "Dificil"
    cmbDificultad.AddItem "Fácil"
    cmbDificultad.AddItem "Muy Fácil"
    cmbDificultad.ListIndex = 2
    
    BalanceFiles = Dir(PathBalance, vbArchive)
    Do While BalanceFiles <> ""
        cmbBalance.AddItem GetVar(PathBalance & BalanceFiles, "INIT", "Nombre")
        BalanceFiles = Dir()
    Loop
    cmbBalance.ListIndex = 0
    
    chkRandom.value = RandomBots
    chkResu.value = SinResu
    chkResuVida.value = ResuNoVida
End Sub

Private Sub imgCruz_Click()
    If frmMain.Visible = False Then
        frmMain.Socket1.Disconnect
        Call Audio.StopWave
        Call LimpiaWsApi
        frmLogin.Show
    End If
    Unload Me
End Sub

Private Sub imgPelear_Click()
    Dificultad = cmbDificultad.ListIndex + 1
    If frmMain.Visible = False Then
        If lstTeam1.ListCount > 0 Then
            If UserCharIndex = 0 Then
                ReDim Preserve Team1(1 To lstTeam1.ListCount + 1) As Integer
            Else
                ReDim Preserve Team1(1 To lstTeam1.ListCount) As Integer
            End If
        End If
        If lstTeam2.ListCount > 0 Then
            ReDim Team2(1 To lstTeam2.ListCount) As Integer
        Else
            ReDim Team2(0 To 0) As Integer
        End If
    End If
    RandomBots = chkRandom.value
    RandomAiBots = chkRandomAi.value
    SinResu = chkResu.value
    ResuNoVida = chkResuVida.value

    If frmMain.Visible = False Then
        frmPersonaje.Modo = 1
        frmPersonaje.Show
    Else
        Call ResetDuelo(-1, lstTeam1.ListCount, lstTeam2.ListCount)
        Unload Me
    End If
End Sub

Private Sub imgPelear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPelear.Picture = LoadPicture(App.path & "\Graficos\Boton_Combatir_Hover.jpg")
End Sub

Private Sub imgSalir_Click()
    Call LimpiaWsApi
    End
End Sub

Private Sub lstTeam1_Click()
    lblClase.Caption = ListaClases(TeamData1(lstTeam1.ListIndex + 1).Clase)
    lblRaza.Caption = ListaRazas(TeamData1(lstTeam1.ListIndex + 1).Raza + 1)
    lblGenero.Caption = IIf(TeamData1(lstTeam1.ListIndex + 1).Genero = 0, "Hombre", "Mujer")
    lblNivel.Caption = TeamData1(lstTeam1.ListIndex + 1).Nivel
    lblTeam.Caption = "Equipo Azul"
    ListSelected = 1
End Sub

Private Sub lstTeam2_Click()
    lblClase.Caption = ListaClases(TeamData2(lstTeam2.ListIndex + 1).Clase)
    lblRaza.Caption = ListaRazas(TeamData2(lstTeam2.ListIndex + 1).Raza + 1)
    lblGenero.Caption = IIf(TeamData2(lstTeam2.ListIndex + 1).Genero = 0, "Hombre", "Mujer")
    lblNivel.Caption = TeamData2(lstTeam2.ListIndex + 1).Nivel
    lblTeam.Caption = "Equipo Rojo"
    ListSelected = 2
End Sub
