VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   0  'None
   Caption         =   "FronBot"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover bot"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar bot"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox lstTeam2 
      BackColor       =   &H8000000A&
      Height          =   1815
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.ListBox lstTeam1 
      BackColor       =   &H8000000A&
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.CheckBox chkResuVida 
      BackColor       =   &H00000000&
      Caption         =   "Check1"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox chkResu 
      BackColor       =   &H00000000&
      Caption         =   "Check1"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   5760
      Width           =   255
   End
   Begin VB.ComboBox cmbBalance 
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CheckBox chkRandom 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   4800
      Width           =   255
   End
   Begin VB.ComboBox cmbDificultad 
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
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblTeam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   18
      Top             =   1320
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblNivel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   16
      Top             =   2280
      Width           =   45
   End
   Begin VB.Label lblGenero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   15
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label lblRaza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   45
   End
   Begin VB.Label lblClase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   13
      Top             =   1560
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2640
      TabIndex        =   7
      Top             =   5760
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   5760
      Width           =   945
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bots Aleatorios:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image imgCruz 
      Height          =   375
      Left            =   4480
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgPelear 
      Height          =   375
      Left            =   1755
      Top             =   3600
      Width           =   1815
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
    
    Me.Picture = LoadPicture(App.path & "\Graficos\configurarbot.jpg")
    
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
    
    chkRandom.Value = RandomBots
    chkResu.Value = SinResu
    chkResuVida.Value = ResuNoVida
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
    RandomBots = chkRandom.Value
    SinResu = chkResu.Value
    ResuNoVida = chkResuVida.Value

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
