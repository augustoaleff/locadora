VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComRelatClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Cadastro"
   ClientHeight    =   3315
   ClientLeft      =   7005
   ClientTop       =   2355
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmbStatus 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox TxtPesquisa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin VB.Frame FrameConsultar 
      Caption         =   "Consultar:"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton OptForn 
         Caption         =   "Fornecedores"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptFunc 
         Caption         =   "Funcionários"
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptClientes 
         Caption         =   "Clientes"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label LblStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label LblPesquisa 
      Caption         =   "Pesquisa:"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Data Cadastro:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label LblA 
      Caption         =   "à"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   135
   End
End
Attribute VB_Name = "FrmComRelatClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TxtPesquisa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        CmbStatus.SetFocus
        
    End If
    
End Sub
Private Sub CmbStatus_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And CmbStatus.Text <> Empty Then
    
        MskPeriodoDe.SetFocus
        
    End If
    
End Sub



Private Sub MskPeriodoDe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskPeriodoDe.Text) <> Empty Or Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = Empty) Then
    
        MskPeriodoAte.SetFocus
        
    End If
    
End Sub

Private Sub MskPeriodoAte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskPeriodoAte.Text) <> Empty Or Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = Empty) Then
    
        CmdImprimir.SetFocus
        
    End If
    
End Sub


