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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   11085
   Begin VB.CommandButton CmdBuscaProduto 
      Caption         =   "Buscar Produto"
      Height          =   495
      Left            =   6000
      TabIndex        =   59
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox TxtDiasAlugados 
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
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdBuscarNome 
      Caption         =   "Buscar por Nome do Cliente"
      Height          =   495
      Left            =   3000
      TabIndex        =   53
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox TxtOBS 
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   10095
   End
   Begin VB.CommandButton CmdRemoverItem 
      Caption         =   "Remover Item"
      Height          =   375
      Left            =   9720
      TabIndex        =   50
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox TxtDesconto 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   49
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox TxtDiferenca 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   47
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox TxtValorRecebido 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   46
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox TxtPagtoMinimo 
      Enabled         =   0   'False
      Height          =   405
      Left            =   9000
      TabIndex        =   43
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox TxtCupomDesconto 
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox TxtCheque 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox TxtCartaoCredito 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox TxtCartaoDebito 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame FrameFormaPagto 
      Caption         =   "Forma de Pagamento"
      Height          =   2775
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   6255
      Begin VB.TextBox TxtQuantCheque 
         Height          =   285
         Left            =   3480
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton CmdValidarCupom 
         Caption         =   "Validar Cupom"
         Height          =   495
         Left            =   3480
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox CmbBandeiraCD 
         Height          =   315
         ItemData        =   "FrmComEmissaoAluguel.frx":0000
         Left            =   3480
         List            =   "FrmComEmissaoAluguel.frx":0010
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox CmbBandeiraCC 
         Height          =   315
         ItemData        =   "FrmComEmissaoAluguel.frx":003D
         Left            =   3480
         List            =   "FrmComEmissaoAluguel.frx":004D
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtParcelasCC 
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TxtDinheiro 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label LblQuantCheque 
         Caption         =   "Quant."
         Height          =   255
         Left            =   2760
         TabIndex        =   51
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label LblBandeiraCC 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LblBandeiraCD 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblParcelasCC 
         Caption         =   "Parcelas"
         Height          =   255
         Left            =   4800
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblCodCupomDesconto 
         Caption         =   "Cód. Cupom de Desconto"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label LblCheque 
         Caption         =   "Cheque"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LblCartaoCredito 
         Caption         =   "Cartão de Crédito"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblCartaoDebito 
         Caption         =   "Cartão de Débito"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblDinheiro 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdCancelarPedido 
      Caption         =   "Cancelar Pedido"
      Height          =   495
      Left            =   240
      TabIndex        =   32
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox TxtValorTotal 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9000
      TabIndex        =   30
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox PctFoto 
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
   End
   Begin VB.TextBox TxtQuant 
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdInserir 
      Caption         =   "Inserir"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TxtCodProduto 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdEmitirPedido 
      Caption         =   "Emitir Pedido"
      Height          =   615
      Left            =   9120
      TabIndex        =   19
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox TxtCodVendedor 
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TxtCodCliente 
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
   Begin VB.Label LblDiasAlugados 
      Caption         =   "Dias Alugados"
      Height          =   255
      Left            =   4320
      TabIndex        =   57
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label LblVendedor 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8400
      TabIndex        =   54
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
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
      Left            =   2400
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
      Height          =   255
      Left            =   7680
      TabIndex        =   48
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label LblValorRecebido 
      Caption         =   "Valor Recebido R$"
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
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label LblCodProduto 
      Caption         =   "Cód. Produto"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label LblDataLimiteDevolucao 
      Caption         =   "Data Limite Devolução"
      Height          =   375
      Left            =   8400
      TabIndex        =   25
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblDataEntrega 
      Caption         =   "Data Entrega"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label LblCodVendedor 
      Caption         =   "Cód.Vendedor"
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
Private Sub CmdBuscaProduto_Click()
 FrmComEmissaoAluguelBuscarProduto.Show
End Sub

Private Sub CmdBuscarNome_Click()
 FrmComEmissaoAluguelBuscarNome.Show
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
 TxtCodVendedor.SetFocus
 
 reg.Close
 
 Else
 
  MsgBox "Emita ou Cancele o Pedido Atual para criar um Novo!", vbExclamation, "Aviso"
 
 End If
End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Top = 100
End Sub

Private Sub MSFlexItens_Click()
Call formata_flex
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
 
 reg.Open ("SELECT DESCRICAO FROM PRODUTOS WHERE CODPROD = " & Trim(TxtCodProduto.Text) & "")
 
 If reg.EOF = False Then
 
 LblProduto.Caption = reg.Fields("Descricao")

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

 TxtDiasAlugados.SetFocus
 

 End If

End Sub
Private Sub TxtDiasAlugados_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxtDiasAlugados.Text <> Empty Then
 
Dim I As Date

I = DateAdd("d", CInt(TxtDiasAlugados.Text), MskDataEntrega.Text)


'I = DateDiff("d", MskDataEntrega.Text, MskDataLimiteDev)

MskDataLimiteDev.Text = I


 TxtOBS.SetFocus

 End If
 
End Sub
Private Sub TxtDinheiro_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then

    TxtCartaoDebito.SetFocus
 End If

End Sub
Private Sub TxtCartaoDebito_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

 If TxtCartaoDebito.Text <> Empty Then
 
    CmbBandeiraCD.SetFocus
    
 Else
    
    TxtCartaoCredito.SetFocus
    
 End If
 
End If

End Sub
Private Sub TxtCartaoCredito_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

 If TxtCartaoCredito.Text <> Empty Then
 
    CmbBandeiraCC.SetFocus
    
 Else
    
    TxtCheque.SetFocus
    
 End If
 
End If

End Sub
Private Sub TxtCheque_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

 If TxtCheque.Text <> Empty Then
 
    TxtQuantCheque.SetFocus
    
 Else
    
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

