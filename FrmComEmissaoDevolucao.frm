VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComEmissaoDevolucao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolução de Filme"
   ClientHeight    =   8400
   ClientLeft      =   3750
   ClientTop       =   2865
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11430
   Begin VB.TextBox TxtDesconto 
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
      Height          =   375
      Left            =   9360
      TabIndex        =   54
      Top             =   6240
      Width           =   1935
   End
   Begin MSMask.MaskEdBox MskDataLimiteDev 
      Height          =   330
      Left            =   5760
      TabIndex        =   53
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSMask.MaskEdBox MskDataEntrega 
      Height          =   330
      Left            =   1440
      TabIndex        =   52
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSMask.MaskEdBox MskDataDev 
      Height          =   330
      Left            =   9480
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.TextBox TxtNumPedido 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox TxtCodVendedor 
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
      Height          =   330
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox TxtCartaoDebito 
      Height          =   285
      Left            =   1800
      TabIndex        =   50
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox TxtCartaoCredito 
      Height          =   285
      Left            =   1800
      TabIndex        =   49
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox TxtCheque 
      Height          =   285
      Left            =   1800
      TabIndex        =   48
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox TxtCupomDesconto 
      Height          =   285
      Left            =   2400
      TabIndex        =   47
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame FrameFormaPagto 
      Caption         =   "Forma de Pagamento"
      Height          =   2775
      Left            =   240
      TabIndex        =   31
      Top             =   4440
      Width           =   6255
      Begin VB.TextBox TxtDinheiro 
         Height          =   285
         Left            =   1560
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtParcelasCC 
         Height          =   285
         Left            =   5520
         TabIndex        =   36
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox CmbBandeiraCC 
         Height          =   315
         ItemData        =   "FrmComEmissaoDevolucao.frx":0000
         Left            =   3480
         List            =   "FrmComEmissaoDevolucao.frx":0002
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox CmbBandeiraCD 
         Height          =   315
         ItemData        =   "FrmComEmissaoDevolucao.frx":0004
         Left            =   3480
         List            =   "FrmComEmissaoDevolucao.frx":0006
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CmdValidarCupom 
         Caption         =   "Validar Cupom"
         Height          =   495
         Left            =   3480
         TabIndex        =   33
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtQuantCheque 
         Height          =   285
         Left            =   3480
         TabIndex        =   32
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label LblDinheiro 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblCartaoDebito 
         Caption         =   "Cartão de Débito"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label LblCartaoCredito 
         Caption         =   "Cartão de Crédito"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblCheque 
         Caption         =   "Cheque"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LblCodCupomDesconto 
         Caption         =   "Cód. Cupom de Desconto"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label LblParcelasCC 
         Caption         =   "Parcelas"
         Height          =   255
         Left            =   4800
         TabIndex        =   41
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LblBandeiraCD 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LblBandeiraCC 
         Caption         =   "Bandeira"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LblQuantCheque 
         Caption         =   "Quant."
         Height          =   255
         Left            =   2760
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.TextBox TxtQuant 
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
      Left            =   8160
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox TxtCodProduto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdBaixarItem 
      Caption         =   "Baixar Item"
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton CmdPesquisa 
      Caption         =   "Pesquisar Pedido por Cliente"
      Height          =   495
      Left            =   4800
      TabIndex        =   24
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton CmdFecharDevolucao 
      Caption         =   "Fechar Devolução"
      Height          =   615
      Left            =   9120
      TabIndex        =   6
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox TxtDiasAtraso 
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
      Left            =   10320
      TabIndex        =   22
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox TxtCodCliente 
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
      Height          =   330
      Left            =   1200
      TabIndex        =   21
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox TxtValorTotal 
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
      Height          =   495
      Left            =   9360
      TabIndex        =   12
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox TxtValorPago 
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
      Height          =   405
      Left            =   9360
      TabIndex        =   11
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox TxtMulta 
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
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox TxtDiferenca 
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
      Height          =   420
      Left            =   9360
      TabIndex        =   9
      Top             =   6840
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexItens 
      Height          =   1815
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.Label LblDesconto 
      Caption         =   "Desconto R$"
      Height          =   375
      Left            =   7920
      TabIndex        =   55
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Pedido"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   240
      Width           =   735
   End
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Left            =   2880
      TabIndex        =   30
      Top             =   720
      Width           =   8295
   End
   Begin VB.Label LblProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3240
      TabIndex        =   29
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label LblQuant 
      Caption         =   "Quant."
      Height          =   255
      Left            =   7440
      TabIndex        =   28
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LblCodProduto 
      Caption         =   "Cód Produto"
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label LblVendedor 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Left            =   5160
      TabIndex        =   26
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label LblCodVendedor 
      Caption         =   "Cód.Vendedor"
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label LblDiasAtraso 
      Caption         =   "Dias de Atraso"
      Height          =   255
      Left            =   9000
      TabIndex        =   20
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label LblDevolucao 
      Caption         =   "Data Limite Devolução"
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label LblDataEntrega 
      Caption         =   "Data Entrega"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód. Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Width           =   855
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
      Left            =   7200
      TabIndex        =   16
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label LblValorPago 
      Caption         =   "Valor Pago R$"
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
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
      Left            =   7560
      TabIndex        =   14
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label LblMulta 
      Caption         =   "Multa/Juros R$"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label LblDataDevolucao 
      Caption         =   "Data Devolução"
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmComEmissaoDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VRECEBIDO, VTOTAL, JUROS, DESCONTO, DIFF As Double
Attribute VTOTAL.VB_VarUserMemId = 1073938432
Attribute JUROS.VB_VarUserMemId = 1073938432
Attribute DESCONTO.VB_VarUserMemId = 1073938432
Attribute DIFF.VB_VarUserMemId = 1073938432
Dim DH, CD, CC, CH As Double
Attribute DH.VB_VarUserMemId = 1073938437
Attribute CD.VB_VarUserMemId = 1073938437
Attribute CC.VB_VarUserMemId = 1073938437
Attribute CH.VB_VarUserMemId = 1073938437
Dim CUPOM As Boolean
Attribute CUPOM.VB_VarUserMemId = 1073938441
Private Sub DIFERENCA()

    DIFF = VTOTAL - VRECEBIDO + JUROS
    TxtDiferenca.Text = Format(DIFF, "#,##0.00")

End Sub
Private Sub carregar_combo_bandeiras()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT Descricao FROM BANDS_CC order by descricao")

    Do Until reg.EOF = True

        CmbBandeiraCC.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

    reg.Open ("SELECT Descricao FROM BANDS_CD order by descricao")

    Do Until reg.EOF = True

        CmbBandeiraCD.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

End Sub
Private Sub limpa_campos()

    TxtCodVendedor.Text = ""
    LblVendedor.Caption = ""
    TxtCodCliente.Text = ""
    LblCliente.Caption = ""
    MskDataEntrega.Mask = ""
    MskDataEntrega.Text = ""
    MskDataEntrega.Mask = "##/##/####"
    MskDataLimiteDev.Mask = ""
    MskDataLimiteDev.Text = ""
    MskDataLimiteDev.Mask = "##/##/####"
    TxtNumPedido.Enabled = True
    MskDataDev.Mask = ""
    MskDataDev.Text = ""
    MskDataDev.Mask = "##/##/####"
    TxtDiasAtraso.Text = ""
    TxtNumPedido.Text = ""
    TxtCodProduto.Text = ""
    LblProduto.Caption = ""
    TxtQuant.Text = ""
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
    TxtValorPago.Text = ""
    TxtMulta.Text = ""
    TxtDesconto.Text = ""
    TxtDiferenca.Text = ""

    VRECEBIDO = 0
    VTOTAL = 0
    DH = 0
    CC = 0
    CD = 0
    CH = 0
    DIFF = 0
    DESCONTO = 0
    JUROS = 0
    CUPOM = False

    Call formata_flex


    TxtNumPedido.SetFocus

End Sub
Private Function VerificaDevolvido() As Boolean

    For contador = 1 To MSFlexItens.Rows - 1

        If MSFlexItens.TextMatrix(contador, 5) <> "ALUGADO" Then

            VerificaDevolvido = True


        Else

            VerificaDevolvido = False
            Exit For

        End If

    Next

End Function
Private Function ValidaCampos() As Boolean

    If TxtNumPedido.Text <> Empty Then

        ValidaCampos = True

        If TxtCodVendedor.Text <> Empty And IsDate(MskDataDev.Text) <> Empty And IsDate(MskDataEntrega.Text) <> Empty And IsDate(MskDataLimiteDev.Text) <> Empty Then

            ValidaCampos = True

            Call VerificaDevolvido

            If VerificaDevolvido = True Then

                ValidaCampos = True

                If CDbl(TxtDiferenca.Text) <= 0# Then

                    ValidaCampos = True

                Else

                    ValidaCampos = False

                    MsgBox "Ainda há Diferença para pagar"

                End If


            Else

                ValidaCampos = False

                MsgBox "Ainda há produtos à serem devolvidos!", vbInformation, Aviso

            End If

        Else

            ValidaCampos = False

            MsgBox "Verifique os Campos", vbInformation, Aviso

        End If

    Else

        ValidaCampos = False

        MsgBox "Verifique os Campos", vbInformation, Aviso

    End If




End Function


Private Sub CmdBaixarItem_Click()

    If TxtCodProduto.Text <> Empty And TxtQuant.Text <> Empty Then


        For contador = 1 To MSFlexItens.Rows - 1

            If MSFlexItens.TextMatrix(contador, 0) = TxtCodProduto.Text Then

                If MSFlexItens.TextMatrix(contador, 5) <> "DEVOLVIDO" Then

                    If MSFlexItens.TextMatrix(contador, 2) = TxtQuant.Text Then


                        MSFlexItens.TextMatrix(contador, 5) = "DEVOLVIDO"
                        TxtCodProduto.Text = ""
                        LblProduto.Caption = ""
                        TxtQuant.Text = ""
                        TxtCodProduto.SetFocus

                    Else

                        MsgBox "A Quantidade devolvida não é a mesma alugada", vbInformation, Aviso


                    End If

                Else

                    MsgBox "Produto já Baixado Anteriormente"

                End If

                'Else

                'MsgBox "Produto não encontrado no pedido", vbInformation, Aviso


            End If

        Next

    Else

        MsgBox "Digite o Produto", vbInformation, Aviso

    End If


End Sub
Private Sub CmdFecharDevolucao_Click()

    Call ValidaCampos

    If ValidaCampos = True Then

        Dim QUANTEST, QUANTALUG As Integer
        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

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



        CN1.Execute ("UPDATE PEDIDOS SET DataDev = '" & Format(MskDataDev.Text, "YYYYMMDD") & "',ValorP = " & Replace(CDbl(VRECEBIDO), ",", ".") & ",Status = 'DEVOLVIDO' " & _
                     "WHERE NumPed = " & TxtNumPedido & "")

        CN1.Execute ("UPDATE ITENS SET Status = 'DEVOLVIDO' WHERE NumPed = " & TxtNumPedido.Text & "")

        For contador = 1 To MSFlexItens.Rows - 1

            reg.Open ("SELECT QuantEst,QuantAlug FROM PRODUTOS WHERE codprod = " & MSFlexItens.TextMatrix(contador, 0) & "")

            If reg.EOF = False Then

                QUANTEST = CInt(reg.Fields("QuantEst"))
                QUANTALUG = CInt(reg.Fields("QuantAlug"))

                QUANTEST = QUANTEST + CInt(MSFlexItens.TextMatrix(contador, 2))
                QUANTALUG = QUANTALUG - CInt(MSFlexItens.TextMatrix(contador, 2))

            End If

            CN1.Execute ("UPDATE PRODUTOS SET QuantEst = '" & CStr(QUANTEST) & "',QuantAlug = '" & CStr(QUANTALUG) & _
                         "' WHERE codprod = " & MSFlexItens.TextMatrix(contador, 0) & "")

            reg.Close
        Next

        If CUPOM = True Then

            CN1.Execute ("INSERT INTO PAGAMENTOS(NumPed,DataPagto,VDinheiro,VCDebito,VCCredito,VCheque,BandCD,BandCC,QuantCC,QuantCH,Juros,Desconto,CodCupom) " & _
                         "VALUES (" & Trim(TxtNumPedido.Text) & ",'" & Format(Now, "YYYYMMDD hh:mm") & "'," & Replace(DH, ",", ".") & "," & Replace(CD, ",", ".") & "," & Replace(CC, ",", ".") & ", " & Replace(CH, ",", ".") & ",'" & StrConv(CmbBandeiraCD.Text, vbUpperCase) & "','" & _
                         StrConv(CmbBandeiraCC.Text, vbUpperCase) & "','" & Trim(TxtParcelasCC.Text) & "','" & Trim(TxtQuantCheque.Text) & "'," & Replace(JUROS, ",", ".") & "," & Replace(DESCONTO, ",", ".") & ",'" & StrConv(Trim(TxtCupomDesconto.Text), vbUpperCase) & "')")

            CN1.Execute ("UPDATE CUPONS SET Status = 'UTILIZADO' " & _
                         "WHERE codCUPOM = '" & StrConv(Trim(TxtCupomDesconto.Text), vbUpperCase) & "'")

        Else


            CN1.Execute ("INSERT INTO PAGAMENTOS(NumPed,DataPagto,VDinheiro,VCDebito,VCCredito,VCheque,BandCD,BandCC,QuantCC,QuantCH,Juros,CodCupom) " & _
                         "VALUES (" & Trim(TxtNumPedido.Text) & ",'" & Format(Now, "YYYYMMDD hh:mm") & "'," & Replace(DH, ",", ".") & "," & Replace(CD, ",", ".") & "," & Replace(CC, ",", ".") & ", " & Replace(CH, ",", ".") & ",'" & StrConv(CmbBandeiraCD.Text, vbUpperCase) & "','" & _
                         StrConv(CmbBandeiraCC.Text, vbUpperCase) & "','" & Trim(CInt(TxtParcelasCC.Text)) & "','" & Trim(CInt(TxtQuantCheque.Text)) & "'," & Replace(JUROS, ",", ".") & ",'')")

        End If


        If TxtDiferenca.Text < 0 Then

            CN1.Execute ("INSERT INTO CUPONS(CodCupom,Tipo,Valor,ValidadeDe,ValidadeAte,Descricao,Status,Usuario,DataEmissao)" & _
                         "VALUES('" & CStr(TxtNumPedido.Text) & "DIF','V'," & Replace(CDbl(Replace(TxtDiferenca.Text, "-", "")), ",", ".") & ",'" & Format(Now, "YYYYMMDD") & "','" & _
                         Format("31/12/2199", "YYYYMMDD") & "','DIF REF PED N " & CStr(TxtNumPedido.Text) & "','NAOUTILIZADO','','" & Format(Now, "YYYYMMDD") & "')")

            CN1.Execute ("UPDATE PEDIDOS SET OBS =  Obs + ' / GERADO CUPOM NUM " & CStr(TxtNumPedido.Text) & "DIF NO VALOR DE R$ " & Replace(Replace(TxtDiferenca.Text, "-", ""), ",", ".") & "' WHERE NUMPED = " & Trim(TxtNumPedido.Text) & "")



            MsgBox "Foi Criado um Cupom no valor de R$ " & Replace(TxtDiferenca.Text, "-", "") & " para a sua proxima locação, Cod Cupom : " & TxtNumPedido.Text & "DIF", vbInformation, "Cupom Criado"

        End If

        'Inserir no Contas a Receber

        If DH <> Empty And DH <> 0 Then

            VENCIMENTO = Format(Now, "YYYY/MM/DD")

            reg.Open ("SELECT CodCli,Vencto,Seq FROM C_A_R WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " AND VENCTO = '" & Format(VENCIMENTO, "YYYYMMDD") & "' ORDER BY SEQ DESC")

            If reg.EOF = True Then


                CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                             "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "',1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                             Trim(TxtNumPedido.Text) & "-DH-D','DH','P','" & Format(Now, "YYYYMMDD hh:mm") & "',''," & Replace(DH, ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")

            Else

                CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                             "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "'," & reg.Fields("Seq") & " + 1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                             Trim(TxtNumPedido.Text) & "-DH-D','DH','P','" & Format(Now, "YYYYMMDD hh:mm") & "',''," & Replace(DH, ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")
            End If

            reg.Close

        End If

        If CD <> Empty And CD <> 0 Then

            VENCIMENTO = Format(DateAdd("d", 1, Now), "YYYY/MM/DD")

            reg.Open ("SELECT CodCli,Vencto,Seq FROM C_A_R WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " AND VENCTO = '" & Format(VENCIMENTO, "YYYYMMDD") & "' order by DataEmissao desc")

            If reg.EOF = True Then

                CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                             "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "',1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                             Trim(TxtNumPedido.Text) & "-CD-D','CD','P','" & Format(Now + 1, "YYYYMMDD hh:mm") & "','BAND: " & CmbBandeiraCD.Text & "'," & Replace(CD, ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")

            Else

                CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                             "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "'," & reg.Fields("Seq") & " + 1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                             Trim(TxtNumPedido.Text) & "-CD-D','CD','P','" & Format(Now + 1, "YYYYMMDD hh:mm") & "','BAND: " & CmbBandeiraCD.Text & "'," & Replace(CD, ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")

            End If

            reg.Close

        End If

        If CC <> Empty And CC <> 0 Then

            VENCIMENTO = Format(Now, "YYYY/MM/DD")

            For contador = 1 To CInt(TxtParcelasCC.Text)


                VENCIMENTO = DateAdd("d", 30, VENCIMENTO)

                reg.Open ("SELECT CodCli,Vencto,Seq FROM C_A_R WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " AND VENCTO = '" & Format(VENCIMENTO, "YYYYMMDD") & "' order by DataEmissao desc")

                If reg.EOF = True Then

                    CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                                 "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "',1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                                 Trim(TxtNumPedido.Text) & "-CC-D-" & contador & "/" & TxtParcelasCC.Text & "','CC','P','" & Format(VENCIMENTO, "YYYYMMDD") & "','BAND: " & CmbBandeiraCC.Text & ", PARCELA : " & contador & "/" & TxtParcelasCC.Text & "'," & Replace(CC / CInt(TxtParcelasCC.Text), ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")

                Else

                    CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                                 "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "'," & reg.Fields("Seq") & " + 1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                                 Trim(TxtNumPedido.Text) & "-CC-D-" & contador & "/" & TxtParcelasCC.Text & "','CC','P','" & Format(VENCIMENTO, "YYYYMMDD") & "','BAND: " & CmbBandeiraCC.Text & ", PARCELA : " & contador & "/" & TxtParcelasCC.Text & "'," & Replace(CC / CInt(TxtParcelasCC.Text), ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")

                End If

                reg.Close


            Next

        End If

        If CH <> Empty And CH <> 0 Then

            VENCIMENTO = Format(Now, "YYYY/MM/DD")

            reg.Open ("SELECT CodCli,Vencto,Seq FROM C_A_R WHERE CODCLI = " & Trim(TxtCodCliente.Text) & " AND VENCTO = '" & Format(VENCIMENTO, "YYYYMMDD") & "' order by DataEmissao desc")

            If reg.EOF = True Then

                CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                             "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "',1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                             Trim(TxtNumPedido.Text) & "-CH-D','CHQ','P','" & Format(Now, "YYYYMMDD hh:mm") & "','QUANT. CHQS: " & TxtQuantCheque.Text & "'," & Replace(CH, ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")
            Else

                CN1.Execute ("INSERT INTO C_A_R(CodCli,Vencto,Seq,DataLancto,NumPed,NumDocto,Tipo,Status,DataPagto,OBS,Valor,Usuario,DataEmissao) " & _
                             "VALUES(" & Trim(TxtCodCliente.Text) & ",'" & Format(VENCIMENTO, "YYYYMMDD") & "'," & reg.Fields("Seq") & " + 1,'" & Format(Now, "YYYYMMDD") & "'," & Trim(TxtNumPedido.Text) & ", '" & _
                             Trim(TxtNumPedido.Text) & "-CH-D','CHQ','P','" & Format(Now, "YYYYMMDD hh:mm") & "','QUANT. CHQS: " & TxtQuantCheque.Text & "'," & Replace(CH, ",", ".") & ",'','" & Format(Now, "YYYYMMDD hh:mm") & "')")
            End If

            reg.Close

        End If


        MsgBox "Devolução do Pedido " & TxtNumPedido.Text & " Concluída!", vbInformation, Confimação

        Call limpa_campos





    Else


    End If



End Sub

Private Sub CmdLimparTela_Click()

    Call limpa_campos

End Sub

Private Sub CmdPesquisa_Click()
    FrmComEmissaoDevolucaoBuscarPedido.Show
End Sub

Private Sub Form_Load()
    Me.Left = 1600
    Me.Top = 1600

    Call carregar_combo_bandeiras

End Sub



Private Sub TxtCodProduto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodProduto.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT DESCRICAO,quantest FROM PRODUTOS WHERE CODPROD = " & Trim(TxtCodProduto.Text) & "")

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


Public Sub TxtNumPedido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtNumPedido.Text) <> Empty Then


        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1
        Set REG2 = New ADODB.Recordset
        REG2.ActiveConnection = CN1


        reg.Open ("SELECT P.NumPed,P.CodVend,F.Nome AS Vendedor,P.CodCli,C.Nome,P.DataEntrega,P.DataLimDev,P.ValorP,P.ValorT,P.Status FROM PEDIDOS AS P INNER JOIN CLIENTES AS C ON P.CODCLI = C.CODCLI INNER JOIN FUNCIONARIOS AS F ON P.CODVEND = F.CODFUNC WHERE NUMPED = " & Trim(TxtNumPedido.Text) & "")

        REG2.Open ("SELECT I.CodProd,P.Descricao,I.Quant,P.Preco,I.Status from ITENS AS I INNER JOIN PRODUTOS AS P ON I.CodProd = P.CodProd WHERE I.NUMPED = " & Trim(TxtNumPedido.Text) & " ")




        If reg.EOF = False Then

            If reg.Fields("Status") = "ALUGADO" Then

                TxtCodVendedor.Text = reg.Fields("CodVend")
                LblVendedor.Caption = reg.Fields("Vendedor")
                TxtCodCliente.Text = reg.Fields("CodCli")
                LblCliente.Caption = reg.Fields("Nome")
                MskDataEntrega.Text = reg.Fields("DataEntrega")
                MskDataLimiteDev.Text = reg.Fields("DataLimDev")
                VRECEBIDO = reg.Fields("ValorP")
                VTOTAL = reg.Fields("ValorT")
                TxtNumPedido.Enabled = False
                TxtCodProduto.SetFocus
                MskDataDev.Text = Format(Now, "DD/MM/YYYY")
                TxtValorPago.Text = Format(VRECEBIDO, "#,##0.00")

                I = DateDiff("d", MskDataEntrega.Text, MskDataDev.Text)

                If I < 0 Then

                    TxtDiasAtraso.Text = 0

                Else

                    TxtDiasAtraso.Text = I

                End If

                'If REG2.EOF = False Then

                Call formata_flex

                Do Until REG2.EOF = True

                    MSFlexItens.AddItem (REG2.Fields("CodProd") & vbTab & _
                                         REG2.Fields("Descricao") & vbTab & _
                                         REG2.Fields("Quant") & vbTab & _
                                         Format(REG2.Fields("Preco"), "#,##0.00") & vbTab & _
                                         Format(REG2.Fields("Quant") * REG2.Fields("Preco"), "#,##0.00") & vbTab & _
                                         REG2.Fields("Status"))

                    REG2.MoveNext


                Loop

                TxtValorTotal.Text = Format(VTOTAL, "#,##0.00")

                If TxtDiasAtraso.Text <> 0 Then

                    JUROS = VTOTAL * (CDbl(TxtDiasAtraso.Text) * 2 / 100)

                Else

                    JUROS = 0

                End If

                TxtMulta.Text = Format(JUROS, "#,##0.00")

                Call DIFERENCA


                'End If
            Else

                MsgBox "Pedido não disponível para Devolução", vbInformation, Aviso
                Call limpa_campos

            End If


        Else

            MsgBox "Pedido não encontrado", vbInformation, Aviso

        End If


        reg.Close
        REG2.Close


    End If


End Sub
Private Sub formata_flex()

    MSFlexItens.Clear
    MSFlexItens.Cols = 6
    MSFlexItens.Rows = 1

    MSFlexItens.Col = 0
    MSFlexItens.Text = "Cód."
    MSFlexItens.ColWidth(0) = 700

    MSFlexItens.Col = 1
    MSFlexItens.Text = "Descricao"
    MSFlexItens.ColWidth(1) = 4200

    MSFlexItens.Col = 2
    MSFlexItens.Text = "Quant."
    MSFlexItens.ColWidth(2) = 900

    MSFlexItens.Col = 3
    MSFlexItens.Text = "Valor Uni."
    MSFlexItens.ColWidth(3) = 900

    MSFlexItens.Col = 4
    MSFlexItens.Text = "Valor Total"
    MSFlexItens.ColWidth(4) = 900

    MSFlexItens.Col = 5
    MSFlexItens.Text = "Status"
    MSFlexItens.ColWidth(5) = 1200

End Sub

Private Sub TxtDinheiro_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then

        TxtCartaoDebito.SetFocus

        If TxtDinheiro.Text <> Empty Then

            VRECEBIDO = VRECEBIDO - DH
            DH = Replace(Replace(TxtDinheiro.Text, "R", ""), "$", "")
            VRECEBIDO = VRECEBIDO + DH
            TxtValorPago.Text = Format(VRECEBIDO, "#,##0.00")
            Call DIFERENCA

        Else

            VRECEBIDO = VRECEBIDO - DH
            DH = 0
            TxtValorPago.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorPago.Text
            Call DIFERENCA

        End If

    End If

End Sub
Private Sub TxtCartaoDebito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If TxtCartaoDebito.Text <> Empty Then

            VRECEBIDO = VRECEBIDO - CD
            CD = Replace(Replace(TxtCartaoDebito.Text, "R", ""), "$", "")
            TxtValorPago.Text = Format(VRECEBIDO + CD, "#,##0.00")
            VRECEBIDO = TxtValorPago.Text
            Call DIFERENCA

            CmbBandeiraCD.SetFocus

        Else

            VRECEBIDO = VRECEBIDO - CD
            CD = 0
            TxtValorPago.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorPago.Text
            Call DIFERENCA

            TxtCartaoCredito.SetFocus

        End If

    End If

End Sub
Private Sub TxtCartaoCredito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If TxtCartaoCredito.Text <> Empty Then


            VRECEBIDO = VRECEBIDO - CC
            CC = Replace(Replace(TxtCartaoCredito.Text, "R", ""), "$", "")
            TxtValorPago.Text = Format(VRECEBIDO + CC, "#,##0.00")
            VRECEBIDO = TxtValorPago.Text
            Call DIFERENCA

            CmbBandeiraCC.SetFocus

        Else

            VRECEBIDO = VRECEBIDO - CC
            CC = 0
            TxtValorPago.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorPago.Text
            Call DIFERENCA

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
            TxtValorPago.Text = Format(VRECEBIDO, "#,##0.00")
            Call DIFERENCA


            TxtQuantCheque.SetFocus

        Else


            VRECEBIDO = VRECEBIDO - CH
            CH = 0
            TxtValorPago.Text = Format(VRECEBIDO, "#,##0.00")
            VRECEBIDO = TxtValorPago.Text
            Call DIFERENCA

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

Private Sub TxtQuant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(TxtQuant.Text) <> Empty Then

        CmdBaixarItem.SetFocus

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

            CmdFecharDevolucao.SetFocus

        End If

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



