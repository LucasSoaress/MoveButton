VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Você é uma pessoa bonita?"
   ClientHeight    =   7920
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnIniciar 
      Appearance      =   0  'Flat
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton BtnNao 
      Appearance      =   0  'Flat
      Caption         =   "Não"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton BtnSim 
      Appearance      =   0  'Flat
      Caption         =   "Sim "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label LblContador 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      Caption         =   "Contador"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   345
      Left            =   960
      TabIndex        =   4
      Top             =   7080
      Width           =   1080
   End
   Begin VB.Label LblResposta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   12495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Você é uma pessoa bonita?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   12615
   End
   Begin VB.Menu Sobre 
      Caption         =   "Sobre o programa"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public status As Boolean
Public contador As Integer
Private Sub BtnIniciar_Click()
Form_Load
status = True
contador = 0
LblResposta.Caption = ""
BtnSim.Enabled = True
BtnSim.Left = 480
BtnNao.Enabled = True
LblContador.Caption = ""
End Sub
Private Sub BtnNao_Click()
    If contador < 1 Then
        LblResposta.Caption = "Tente usar o sim, só uma vez"
        Else
            LblResposta.Caption = "Você demorou " & contador & " tentativas para admitir que é uma pessoa feia!!! hahahaha"
    End If
BtnSim.Enabled = False
BtnNao.Enabled = False
End Sub
Private Sub BtnSim_Click()
MsgBox "Parabéns você realmente é uma pessoa bonita, e superou o desafio do programa"
End Sub
Private Sub BtnSim_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If status = True Then
BtnSim.Left = 3840
status = False
    Else
BtnSim.Left = 480
status = True
    End If
contador = contador + 1
LblContador.Caption = contador
End Sub
Private Sub Form_Load()
status = True
contador = 0
End Sub
Private Sub Sobre_Click()
 Form2.Show
 End Sub
