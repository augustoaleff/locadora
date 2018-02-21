VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComEmissaoAluguel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Aluguel de Filme"
   ClientHeight    =   9225
   ClientLeft      =   225
   ClientTop       =   2505
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   11085
   Begin VB.CommandButton CmdBuscaProduto 
      Caption         =   "Buscar Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   59
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox TxtDiasAlugados 
      Enabled         =   0   'False
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
      Left            =   5520
      TabIndex        =   58
      Top             =   1440
      Width           =   855
   End
   Begin MSMask.MaskEdBox MskDataEntrega 
      Height          =   330
      Left            =   1680
      TabIndex        =   55
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VB.CommandButton CmdGerarNumeroPedido 
      Caption         =   "Gerar nº Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdBuscarNome 
      Caption         =   "Buscar por Nome do Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   53
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox TxtOBS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   10095
   End
   Begin VB.CommandButton CmdRemoverItem 
      Caption         =   "Remover Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   50
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TxtDesconto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   49
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox TxtDiferenca 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   47
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox TxtValorRecebido 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   46
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox TxtPagtoMinimo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   43
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox TxtCupomDesconto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox TxtCheque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox TxtCartaoCredito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox TxtCartaoDebito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame FrameFormaPagto 
      Caption         =   "Forma de Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   6255
      Begin VB.TextBox TxtQuantCheque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton CmdValidarCupom 
         Caption         =   "Validar Cupom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox CmbBandeiraCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmComEmissaoAluguel.frx":0000
         Left            =   3480
         List            =   "FrmComEmissaoAluguel.frx":0010
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox CmbBandeiraCC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmComEmissaoAluguel.frx":003D
         Left            =   3480
         List            =   "FrmComEmissaoAluguel.frx":004D
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtParcelasCC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TxtDinheiro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label LblQuantCheque 
         Caption         =   "Quant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   51
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label LblBandeiraCC 
         Caption         =   "Bandeira"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LblBandeiraCD 
         Caption         =   "Bandeira"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblParcelasCC 
         Caption         =   "Parcelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblCodCupomDesconto 
         Caption         =   "Cód. Cupom de Desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label LblCheque 
         Caption         =   "Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LblCartaoCredito 
         Caption         =   "Cartão de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblCartaoDebito 
         Caption         =   "Cartão de Débito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblDinheiro 
         Caption         =   "Dinheiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdCancelarPedido 
      Caption         =   "Cancelar Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   32
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox TxtValorTotal 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   30
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox PctFoto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9120
      ScaleHeight     =   1755
      ScaleWidth      =   1635
      TabIndex        =   29
      Top             =   3240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexItens 
      Height          =   1935
      Left            =   240
      TabIndex        =   28
      Top             =   3240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3413
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtQuant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton CmdInserir 
      Caption         =   "Inserir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TxtCodProduto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdEmitirPedido 
      Caption         =   "Emitir Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   19
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox TxtCodVendedor 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TxtCodCliente 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox TxtNumPedido 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MskDataLimiteDev 
      Height          =   330
      Left            =   9480
      TabIndex        =   56
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label LblQuantEst 
      Caption         =   "Quant.Est"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   61
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label LblQuantEstoque 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   60
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label LblDiasAlugados 
      Caption         =   "Dias Alugados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   57
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label LblVendedor 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   54
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label LblProduto 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   840
      Width           =   8055
   End
   Begin VB.Label LblDesconto 
      Caption         =   "Desconto R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   48
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label LblValorRecebido 
      Caption         =   "Valor Recebido R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   45
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label LblDiferenca 
      Caption         =   "Diferença R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   44
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label LblPgtoMinimo 
      Caption         =   "Pagamento Mínimo R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   37
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label LblValorTotal 
      Caption         =   "Valor Total R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   31
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label LblQuantProduto 
      Caption         =   "Quant."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   27
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label LblCodProduto 
      Caption         =   "Cód. Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label LblDataLimiteDevolucao 
      Caption         =   "Data Limite Devolução"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   25
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblDataEntrega 
      Caption         =   "Data Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label LblCodVendedor 
      Caption         =   "Cód.Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   23
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód. Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label LblNumPedido 
      Caption         =   "Pedido nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmComEmissaoAluguel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VRECEBIDO, VTOTAL, DH, CD, CC, CH As Double
Attribute VTOTAL.VB_VarUserMemId = 1073938432
Attribute DH.VB_VarUserMemId = 1073938432
Attribute CD.VB_VarUserMemId = 1073938432
Attribute CC.VB_VarUserMemId = 1073938432
Attribute CH.VB_VarUserMemId = 1073938432
Dim PAGTOMIN, DIFERENCA, DESCONTO As Double
Attribute PAGTOMIN.VB_VarUserMemId = 1073938438
Attribute DIFERENCA.VB_VarUserMemId = 1073938438
Attribute DESCONTO.VB_VarUserMemId = 1073938438
Dim CUPOM As Boolean
Attribute CUPOM.VB_VarUserMemId = 1073938441
Private Sub calcula_diferenca()

    DIFERENCA = VTOTAL - VRECEBIDO
    TxtDiferenca.Text = Format(DIFERENCA, "#,##0.00")

End Sub
Private Function verifica_produto() As Boolean

    If MSFlexItens.Rows >= 2 Then

        For contador = 1 To MSFlexItens.Rows - 1

            If MSFlexItens.TextMatrix(contador, 0) <> Trim(TxtCodProduto.Text) Then

                verifica_produto = True

            Else

                verifica_produto = False

                Exit For

            End If

        Next

    Else

        verifica_produto = True

    End If


End Function

Private Sub CmdBuscaProduto_Click()
    FrmComEmissaoAluguelBuscarProduto.Show
End Sub

Private Sub CmdBuscarNome_Click()
    FrmComEmissaoAluguelBuscarNome.Show
End Sub

Private Function ValidaCampos() As Boolean

    If IsNumeric(TxtNumPedido.Text) <> Empty Then

        ValidaCampos = True

        If IsNumeric(TxtCodVendedor.Text) <> Empty Then

            ValidaCampos = True

            If IsNumeric(TxtCodCliente.Text) <> Empty Then

                ValidaCampos = True

                If IsDate(MskDataEntrega.Text) <> Empty Then

                    ValidaCampos = True

                    If TxtDiasAlugados.Text <> Empty Then

                        ValidaCampos = True

                        If IsDate(MskDataLimiteDev.Text) <> Empty Then

                            ValidaCampos = True

                            If VTOTAL <> Empty And VTOTAL <> 0 Then

                                ValidaCampos = True

                                If (DIFERENCA - PAGTOMIN) <= 0 Then

                                    ValidaCampos = True

                                    If (TxtCartaoDebito.Text <> Empty And CmbBandeiraCD <> Empty) Or TxtCartaoDebito = Empty Then

                                        ValidaCampos = True

                                        If (TxtCartaoCredito.Text <> Empty And CmbBandeiraCC <> Empty And TxtParcelasCC.Text <> Empty) Or TxtCartaoCredito = Empty Then

                                            ValidaCampos = True

                                            If (TxtCheque.Text <> Empty And TxtQuantCheque <> Empty) Or TxtCheque = Empty Then

                                                ValidaCampos = True

                                            Else

                                                ValidaCampos = False

                                                MsgBox "Verifique o campo de Cheque", vbExclamation, Atenção

                                            End If

                                        Else

                                            ValidaCampos = False

                                            MsgBox "Verifique o campo de Cartão de Crédito", vbExclamation, Atenção

                                        End If

                                    Else

                                        ValidaCampos = False

                                        MsgBox "Verifique o campo de Cartão de Débito", vbExclamation, Atenção

                                    End If

                                Else

                                    ValidaCampos = False

                                    MsgBox "O Pagamento Mínimo não foi alcançado!", vbExclamation, Atenção

                                End If

                            Else

                                ValidaCampos = False

                            End If

                        Else

                            ValidaCampos = False

                        End If

                    Else

                        ValidaCampos = False

                    End If

                Else

                    ValidaCampos = False

                End If

            Else

                ValidaCampos = False

            End If

        Else

            ValidaCampos = False

        End If

    Else

        ValidaCampos = False

    End If


End Function
Private Sub limpa_campos()

    TxtNumPedido.Text = ""
    TxtNumPedido.Enabled = False
    TxtCodVendedor.Text = ""
    TxtCodVendedor.Enabled = False
    LblVendedor.Caption = ""
    TxtCodCliente.Text = ""
    TxtCodCliente.Enabled = False
    LblCliente.Caption = ""
    MskDataEntrega.Mask = ""
    MskDataEntrega.Text = ""
    MskDataEntrega.Mask = "##/##/####"
    MskDataEntrega.Enabled = False
    TxtDiasAlugados.Text = ""
    TxtDiasAlugados.Enabled = False
    MskDataLimiteDev.Mask = ""
    MskDataLimiteDev.Text = ""
    MskDataLimiteDev.Mask = "##/##/####"
    MskDataLimiteDev.Enabled = False
    TxtCodProduto.Text = ""
    LblProduto.Caption = ""
    TxtQuant.Text = ""
    TxtOBS.Text = ""
    TxtDinheiro.Text = ""
    TxtCartaoCredito.Text = ""
    TxtCartaoDebito.Text = ""
    TxtCheque.Text = ""
    CmbBandeiraCC.Text = ""
    CmbBandeiraCD.Text = ""
    TxtParcelasCC.Text = ""
    TxtQuantCheque.Text = ""
    TxtCupomDesconto.Text = ""
    TxtValorTotal.Text = ""
    TxtPagtoMinimo.Text = ""
    TxtValorRecebido.Text = ""
    TxtDesconto.Text = ""
    TxtDiferenca.Text = ""
    LblQuantEstoque.Caption = ""

    VRECEBIDO = 0
    VTOTAL = 0
    DH = 0
    CD = 0
    CC = 0
    CH = 0
    PAGTOMIN = 0
    DESCONTO = 0
    DIFERENCA = 0
    CUPOM = False

    Call formata_flex

End Sub

Private Sub CmdCancelarPedido_Click()
    Dim QUANTEST, QUANTALUG As Integer
    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    resultado = MsgBox("Deseja Cancelar o Pedido", vbYesNo, Confirmação)

    If resultado = vbYes Then

        For contador = 1 To MSFlexItens.Rows - 1

            reg.Open ("SELECT QuantEst,QuantAlug FROM PRODUTOS WHERE codprod = " & MSFlexItens.TextMatrix(contador, 0) & "")

            If reg.EOF = False Then

                QUANTEST = CInt(reg.Fields("QuantEst"))
                QUANTALUG = CInt(reg.Fields("QuantAlug"))

                QUANTEST = QUANTEST + CInt(MSFlexItens.TextMatrix(contador, 2))
                QUANTALUG = QUANTALUG - CInt(MSFlexItens.TextMatrix(contador, 2))



                CN1.Execute ("UPDATE PRODUTOS SET QuantEst = '" & CStr(QUANTEST) & "',QuantAlug = '" & CStr(QUANTALUG) & _
                             "' WHERE codprod = " & MSFlexItens.TextMatrix(contador, 0) & "")


            End If


            reg.Close
        Next



        Call limpa_campos
        MsgBox "Pedido Cancelado!", vbExclamation, Aviso

    End If

End Sub

Private Sub CmdEmitirPedido_Click()


    resultado = MsgBox("Confima a emissao do pedido?", vbYesNo, Confirmação)

    If resultado = vbYes Then

        If ValidaCampos = True Then


            Set CN1 = New ADODB.Connection
            CN1.Open STR_DSN

            If CC = Empty Then
                CC = 0
            End If

            If DH = Empty Then
                DH = 0
            End If

            If CD = Empty Then
                CD = 0
            End If

            If CH = Empty Then
                CH = 0
            End If

            If TxtParcelasCC.Text = Empty Then
                TxtParcelasCC.Text = "0"
            End If

            If TxtQuantCheque.Text = Empty Then
                TxtQuantCheque.Text = "0"
            End If

            If DESCONTO = Empty Then
                DESCONTO = 0
            End If



            CN1.Execute ("INSERT INTO PEDIDOS(NumPed,CodVend,CodCli,DataEntrega,DataLimDev,DataDev,OBS,ValorT,ValorP,Status,Usuario,DataEmissao)" & _
                         "VALUES(" & Trim(CInt(TxtNumPedido.Text)) & "," & Trim(CInt(TxtCodVendedor.Text)) & "," & Trim(CInt(TxtCodCliente.Text)) & ",'" & Format(MskDataEntrega.Text, "YYYYMMDD") & "','" & _
                         Format(MskDataLimiteDev.Text, "YYYYMMDD") & "','','" & Trim(StrConv(TxtOBS.Text, vbUpperCase)) & "'," & Replace(VTOTAL, ",", ".") & "," & Replace(VRECEBIDO, ",", ".") & ",'ALUGADO','','" & Format(Now, "YYYYMMDD hh:mm") & "')")

            For contador = 1 To MSFlexItens.Rows - 1

                CN1.Execute ("INSERT INTO ITENS(NumPed,CodProd,Quant,Status)" & _
                             "VALUES(" & Trim(TxtNumPedido.Text) & "," & MSFlexItens.TextMatrix(contador, 0) & ",'" & MSFlexItens.TextMatrix(contador, 2) & "','ALUGADO')")

            Next

            If CUPOM = True Then

                CN1.Execute ("INSERT INTO PAGAMENTOS(NumPed,DataPagto,VDinheiro,VCDebito,VCCredito,VCheque,BandCD,BandCC,QuantCC,QuantCH,Juros,Desconto,CodCupom) " & _
                             "VALUES (" & Trim(TxtNumPedido.Text) & ",'" & Format(Now, "YYYYMMDD hh:mm") & "'," & Replace(DH, ",", ".") & "," & Replace(CD, ",", ".") & "," & Replace(CC, ",", ".") & ", " & Replace(CH, ",", ".") & ",'" & StrConv(CmbBandeiraCD.Text, vbUpperCase) & "','" & _
                             StrConv(CmbBandeiraCC.Text, vbUpperCase) & "','" & Trim(CInt(TxtParcelasCC.Text)) & "','" & Trim(CInt(TxtQuantCheque.Text)) & "',''," & Replace(DESCONTO, ",", ".") & ",'" & StrConv(Trim(TxtCupomDesconto.Text), vbUpperCase) & "')")

                CN1.Execute ("UPDATE CUPONS SET Status = 'UTILIZADO' " & _
                             "WHERE codCUPOM = '" & StrConv(Trim(TxtCupomDesconto.Text), vbUpperCase) & "'")

            Else


                CN1.Execute ("INSERT INTO PAGAMENTOS(NumPed,DataPagto,VDinheiro,VCDebito,VCCredito,VCheque,BandCD,BandCC,QuantCC,QuantCH,Juros,CodCupom) " & _
                             "VALUES (" & Trim(TxtNumPedido.Text) & ",'" & Format(Now, "YYYYMMDD hh:mm") & "'," & Replace(DH, ",", ".") & "," & Replace(CD, ",", ".") & "," & Replace(CC, ",", ".") & ", " & Replace(CH, ",", ".") & ",'" & StrConv(CmbBandeiraCD.Text, vbUpperCase) & "','" & _
                             StrConv(CmbBandeiraCC.Text, vbUpperCase) & "','" & Trim(CInt(TxtParcelasCC.Text)) & "','" & Trim(CInt(TxtQuantCheque.Text)) & "','','')")

            End If

            MsgBox "Pedido " & TxtNumPedido.Text & " Emitido", vbInformation, Confimação

            Call limpa_campos

        Else

            MsgBox "Verifique os campos!", vbInformation, Aviso

        End If


    End If
End Sub

Private Sub CmdGerarNumeroPedido_Click()

    If TxtNumPedido.Text = Empty Then

        TxtNumPedido.Enabled = True

        Dim QUERY As String

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1


        CN1.Execute ("begin transaction")
        QUERY = "select UltNumPedido from parametros WITH (ROWLOCK);UPDATE PARAMETROS WITH(ROWLOCK) SET UltNumPedido = UltNumPedido+7;COMMIT"
        reg.Open (QUERY)

        TxtNumPedido.Text = reg.Fields("UltNumPedido")
        TxtNumPedido.Enabled = False
        TxtCodVendedor.Enabled = True
        TxtCodVendedor.SetFocus

        reg.Close

    Else

        MsgBox "Emita ou Cancele o Pedido Atual para criar um Novo!", vbExclamation, "Aviso"

    End If
End Sub

Private Sub CmdInserir_Click()

    If verifica_produto = True Then


        Dim QUANTEST, QUANTALUG As Integer
        Dim TOTAL As Double
        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set REG2 = New ADODB.Recordset
        REG2.ActiveConnection = CN1
        Set REG3 = New ADODB.Recordset
        REG3.ActiveConnection = CN1


        REG3.Open ("SELECT QuantEst FROM PRODUTOS WHERE codprod = " & Trim(TxtCodProduto.Text) & "")

        If CInt(REG3.Fields("QuantEst")) >= CInt(TxtQuant.Text) Then

            REG3.Close

            REG2.Open ("SELECT codprod,descricao,preco,QuantEst,QuantAlug FROM PRODUTOS WHERE codprod = " & Trim(TxtCodProduto.Text) & "")

            TOTAL = REG2.Fields("preco") * TxtQuant.Text

            MSFlexItens.AddItem (REG2.Fields("CodProd") & vbTab & _
                                 REG2.Fields("Descricao") & vbTab & _
                                 TxtQuant.Text & vbTab & _
                                 Format(REG2.Fields("preco"), "#,##0.00") & vbTab & _
                                 Format(TOTAL, "#,##0.00"))

            If REG2.EOF = False Then

                QUANTEST = CInt(REG2.Fields("QuantEst"))
                QUANTALUG = CInt(REG2.Fields("QuantAlug"))

                QUANTEST = QUANTEST - CInt(TxtQuant.Text)
                QUANTALUG = QUANTALUG + CInt(TxtQuant.Text)



                CN1.Execute ("UPDATE PRODUTOS SET QuantEst = '" & CStr(QUANTEST) & "',QuantAlug = '" & CStr(QUANTALUG) & _
                             "' WHERE codprod = " & Trim(TxtCodProduto.Text) & "")


            End If

            REG2.Close

            TxtCodProduto.Text = ""
            LblProduto.Caption = ""
            TxtQuant.Text = ""
            LblQuantEstoque.Caption = ""
            TxtCodProduto.SetFocus

            VTOTAL = 0

            For contador2 = 1 To MSFlexItens.Rows - 1
                MSFlexItens.Row = contador2
                MSFlexItens.Col = 4
                VTOTAL = VTOTAL + CDbl(MSFlexItens.Text)

            Next

            TxtValorTotal.Text = Format(VTOTAL, "#,##0.00")
            PAGTOMIN = VTOTAL * 0.5
            TxtPagtoMinimo.Text = Format(PAGTOMIN, "#,##0.00")


        Else

            MsgBox "Estoque menor que o solicitado!", vbInformation, Aviso

        End If
    Else

        MsgBox "Produto já inserido anteriormente", vbInformation, Aviso

    End If


End Sub

Private Sub CmdRemoverItem_Click()


    Dim QUANTEST, QUANTALUG As Integer
    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    If MSFlexItens.Rows = 2 Or MSFlexItens.Rows = 1 Then

        Call formata_flex

        VTOTAL = 0

        TxtValorTotal.Text = Format(VTOTAL, "#,##0.00")
        PAGTOMIN = VTOTAL * 0.5
        TxtPagtoMinimo.Text = Format(PAGTOMIN, "#,##0.00")

        reg.Open ("SELECT QuantEst,QuantAlug FROM PRODUTOS WHERE codprod = " & MSFlexItens.TextMatrix(contador, 0) & "")

        If reg.EOF = False Then

            QUANTEST = CInt(reg.Fields("QuantEst"))
            QUANTALUG = CInt(reg.Fields("QuantAlug"))

            QUANTEST = QUANTEST - CInt(MSFlexItens.TextMatrix(1, 2))
            QUANTALUG = QUANTALUG + CInt(MSFlexItens.TextMatrix(1, 2))



            CN1.Execute ("UPDATE PRODUTOS SET QuantEst = '" & CStr(QUANTEST) & "',QuantAlug = '" & CStr(QUANTALUG) & _
                         "' WHERE codprod = " & MSFlexItens.TextMatrix(1, 0) & "")


        End If


        reg.Close


    Else

        MSFlexItens.RemoveItem (MSFlexItens.RowSel)

        reg.Open ("SELECT QuantEst,QuantAlug FROM PRODUTOS WHERE codprod = " & MSFlexItens.TextMatrix(MSFlexItens.RowSel, 0) & "")

        If reg.EOF = False Then

            QUANTEST = CInt(reg.Fields("QuantEst"))
            QUANTALUG = CInt(reg.Fields("QuantAlug"))

            QUANTEST = QUANTEST + CInt(MSFlexItens.TextMatrix(MSFlexItens.RowSel, 2))
            QUANTALUG = QUANTALUG - CInt(MSFlexItens.TextMatrix(MSFlexItens.RowSel, 2))



            CN1.Execute ("UPDATE PRODUTOS SET QuantEst = '" & CStr(QUANTEST) & "',QuantAlug = '" & CStr(QUANTALUG) & _
                         "' WHERE codprod = " & MSFlexItens.TextMatrix(MSFlexItens.RowSel, 0) & "")


        End If


        reg.Close
        VTOTAL = 0

        For contador2 = 1 To MSFlexItens.Rows - 1
            MSFlexItens.Row = contador2
            MSFlexItens.Col = 4
            VTOTAL = VTOTAL + CDbl(MSFlexItens.Text)
        Next

        TxtValorTotal.Text = Format(VTOTAL, "#,##0.00")
        PAGTOMIN = VTOTAL * 0.5
        TxtPagtoMinimo.Text = Format(PAGTOMIN, "#,##0.00")

    End If

End Sub

Private Sub CmdValidarCupom_Click()

    CUPOM = False
    Dim I, F As Integer

    If TxtCupomDesconto.Text <> Empty Then


        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1


        reg.Open ("SELECT * FROM CUPONS WHERE CODCUPOM = '" & TxtCupomDesconto.Text & "'")



        If reg.EOF = False Then

            I = DateDiff("d", Now, reg.Fields("ValidadeAte"))
            F = DateDiff("d", reg.Fields("ValidadeDe"), Now)

            If I >= 0 And F >= 0 Then

                If reg.Fields("Tipo") = "V" And reg.Fields("Status") = "NAOUTILIZADO" Then


                    VRECEBIDO = VRECEBIDO - DESCONTO
                    DESCONTO = CDbl(reg.Fields("Valor"))
                    MsgBox "Cupom Validado", vbInformation, CUPOM
                    TxtDesconto.Text = Format(DESCONTO, "#,##0.00")
                    VRECEBIDO = VRECEBIDO + DESCONTO
                    TxtValorRecebido = Format(VRECEBIDO, "#,##0.00")
                    Call calcula_diferenca
                    CUPOM = True


                ElseIf reg.Fields("Tipo") = "P" And reg.Fields("Status") = "NAOUTILIZADO" Then

                    VRECEBIDO = VRECEBIDO - DESCONTO
                    DESCONTO = VTOTAL * CDbl(reg.Fields("Valor"))
                    MsgBox "Cupom Validado", vbInformation, CUPOM
                    TxtDesconto.Text = Format(DESCONTO, "#,##0.00")
                    VRECEBIDO = VRECEBIDO + DESCONTO
                    TxtValorRecebido = Format(VRECEBIDO, "#,##0.00")
                    Call calcula_diferenca
                    CUPOM = True

                Else

                    MsgBox "Cupom já Utilizado", vbInformation, Aviso

                End If



            Else

                MsgBox "O Período do Cupom é de " & Format(reg.Fields("ValidadeDe"), "DD/MM/YYYY") & " até " & Format(reg.Fields("ValidadeAte"), "DD/MM/YYYY") & " !", vbExclamation, Aviso

            End If

        Else

            MsgBox "Cupom não encontrado", vbExclamation, Aviso
            CUPOM = False

        End If


    Else

        MsgBox "Digite o cupom", vbExclamation, Aviso
        CUPOM = False
    End If


End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Top = 100

    Call formata_flex

    CUPOM = False
End Sub

Public Sub TxtCodCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodCliente.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT Nome FROM CLIENTES WHERE CODCLI = " & Trim(TxtCodCliente.Text) & "")

        If reg.EOF = False Then

            LblCliente.Caption = reg.Fields("Nome")

            MskDataEntrega.Enabled = True
            MskDataEntrega.SetFocus

        Else

            MsgBox "Cliente Não Encontrado", vbExclamation, "Aviso"
            TxtCodCliente.SetFocus

        End If

        reg.Close


    End If

End Sub
Public Sub TxtCodProduto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodProduto.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT DESCRICAO,quantest FROM PRODUTOS WHERE CODPROD = " & Trim(TxtCodProduto.Text) & "")

        If reg.EOF = False Then

            LblProduto.Caption = reg.Fields("Descricao")
            LblQuantEstoque.Caption = reg.Fields("QuantEst")

            TxtQuant.SetFocus

        Else

            MsgBox "Produto Não Encontrado", vbExclamation, "Aviso"
            TxtCodProduto.SetFocus

        End If

        reg.Close


    End If

End Sub

Private Sub TxtCodVendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodVendedor.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT Nome FROM FUNCIONARIOS WHERE CODFUNC = " & Trim(TxtCodVendedor.Text) & " AND CARGO LIKE '%VEND%'")

        If reg.EOF = False Then

            LblVendedor.Caption = reg.Fields("Nome")

            TxtCodCliente.Enabled = True
            TxtCodCliente.SetFocus

        Else

            MsgBox "Vendedor Não Existe ou não é Vendedor", vbExclamation, "Aviso"
            TxtCodVendedor.SetFocus

        End If

        reg.Close


    End If

End Sub
Private Sub MskDataEntrega_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsDate(MskDataEntrega.Text) <> Empty Then

        TxtDiasAlugados.Enabled = True
        TxtDiasAlugados.SetFocus


    End If

End Sub

Private Sub TxtDiasAlugados_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtDiasAlugados.Text <> Empty Then

        Dim I As Date

        I = DateAdd("d", CInt(TxtDiasAlugados.Text), MskDataEntrega.Text)


        'I = DateDiff("d", MskDataEntrega.Text, MskDataLimiteDev)

        MskDataLimiteDev.Text = I


        TxtOBS.Enabled = True
        TxtOBS.SetFocus


    End If

End Sub
Private Sub TxtDinheiro_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then

        TxtCartaoDebito.SetFocus

        If TxtDinheiro.Text <> Empty Then

            VRECEBIDO = VRECEBIDO - DH
            DH = Replace(Replace(TxtDinheiro.Text, "R", ""), "$", "")
            VRECEBIDO = VRECEBIDO + DH
            TxtValorRecebido.Text = Format(VRECEBIDO, "#,##0.00")
            Call calcula_diferenca

        Else

            VRECEBIDO = VRECEBIDO - DH
            DH = 0
            TxtValorRecebido.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorRecebido.Text
            Call calcula_diferenca

        End If

    End If

End Sub
Private Sub TxtCartaoDebito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If TxtCartaoDebito.Text <> Empty Then

            VRECEBIDO = VRECEBIDO - CD
            CD = Replace(Replace(TxtCartaoDebito.Text, "R", ""), "$", "")
            TxtValorRecebido.Text = Format(VRECEBIDO + CD, "#,##0.00")
            VRECEBIDO = TxtValorRecebido.Text
            Call calcula_diferenca

            CmbBandeiraCD.SetFocus

        Else

            VRECEBIDO = VRECEBIDO - CD
            CD = 0
            TxtValorRecebido.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorRecebido.Text
            Call calcula_diferenca

            TxtCartaoCredito.SetFocus

        End If

    End If

End Sub
Private Sub TxtCartaoCredito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If TxtCartaoCredito.Text <> Empty Then


            VRECEBIDO = VRECEBIDO - CC
            CC = Replace(Replace(TxtCartaoCredito.Text, "R", ""), "$", "")
            TxtValorRecebido.Text = Format(VRECEBIDO + CC, "#,##0.00")
            VRECEBIDO = TxtValorRecebido.Text
            Call calcula_diferenca

            CmbBandeiraCC.SetFocus

        Else

            VRECEBIDO = VRECEBIDO - CC
            CC = 0
            TxtValorRecebido.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorRecebido.Text
            Call calcula_diferenca

            TxtCheque.SetFocus

        End If

    End If

End Sub
Private Sub TxtCheque_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If TxtCheque.Text <> Empty Then

            VRECEBIDO = VRECEBIDO - CH
            CH = Replace(Replace(TxtCheque.Text, "R", ""), "$", "")
            VRECEBIDO = VRECEBIDO + CH
            TxtValorRecebido.Text = Format(VRECEBIDO, "#,##0.00")
            Call calcula_diferenca


            TxtQuantCheque.SetFocus

        Else


            VRECEBIDO = VRECEBIDO - CH
            CH = 0
            TxtValorRecebido.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorRecebido.Text
            Call calcula_diferenca

            TxtCupomDesconto.SetFocus

        End If

    End If

End Sub

Private Sub CmbBandeiraCD_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And CmbBandeiraCD.Text <> Empty Then

        TxtCartaoCredito.SetFocus

    End If

End Sub
Private Sub CmbBandeiraCC_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And CmbBandeiraCC.Text <> Empty Then

        TxtParcelasCC.SetFocus

    End If

End Sub
Private Sub TxtParcelasCC_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtParcelasCC.Text <> Empty Then

        TxtCheque.SetFocus

    End If

End Sub
Private Sub TxtQuantCheque_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtQuantCheque.Text <> Empty Then

        TxtCupomDesconto.SetFocus

    End If

End Sub
Private Sub TxtCupomDesconto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If TxtCupomDesconto.Text <> Empty Then

            CmdValidarCupom.SetFocus

        Else

            CmdEmitirPedido.SetFocus

        End If

    End If

End Sub

Private Sub TxtOBS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        TxtCodProduto.SetFocus

    End If

End Sub

Private Sub TxtQuant_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (TxtQuant.Text) <> Empty Then

        CmdInserir.SetFocus

    End If


End Sub
Private Sub formata_flex()

    MSFlexItens.Clear
    MSFlexItens.Cols = 5
    MSFlexItens.Rows = 1

    MSFlexItens.Col = 0
    MSFlexItens.Text = "Cód."
    MSFlexItens.ColWidth(0) = 700

    MSFlexItens.Col = 1
    MSFlexItens.Text = "Descricao"
    MSFlexItens.ColWidth(1) = 4600

    MSFlexItens.Col = 2
    MSFlexItens.Text = "Quant."
    MSFlexItens.ColWidth(2) = 900

    MSFlexItens.Col = 3
    MSFlexItens.Text = "Valor Uni."
    MSFlexItens.ColWidth(3) = 900

    MSFlexItens.Col = 4
    MSFlexItens.Text = "Valor Total"
    MSFlexItens.ColWidth(4) = 900

End Sub







