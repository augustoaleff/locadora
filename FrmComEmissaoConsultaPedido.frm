VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComEmissaoConsultaPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Pedido"
   ClientHeight    =   9405
   ClientLeft      =   1695
   ClientTop       =   735
   ClientWidth     =   10335
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   10335
   Begin VB.TextBox TxtJuros 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   46
      Top             =   7800
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MskDataEmissao 
      Height          =   330
      Left            =   8760
      TabIndex        =   41
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   3720
      TabIndex        =   38
      Top             =   8640
      Width           =   2295
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   8520
      Width           =   2775
   End
   Begin VB.CommandButton CmdCancelarPedido 
      Caption         =   "Cancelar Pedido"
      Height          =   375
      Left            =   240
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Frame FrameFormaPagto 
      Caption         =   "Forma de Pagamento"
      Height          =   2655
      Left            =   1680
      TabIndex        =   23
      Top             =   5640
      Width           =   6975
      Begin VB.TextBox TxtDesconto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   34
         Top             =   1800
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexCupons 
         Height          =   1935
         Left            =   3840
         TabIndex        =   33
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
         _Version        =   393216
      End
      Begin VB.TextBox TxtTotalCheque 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   31
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalCC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   30
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalCD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalDinheiro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Total Juros/Multa R$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label LblTotalDesconto 
         Caption         =   "Total Desconto R$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label LblCuponsUtilizados 
         Caption         =   "Cupons Utilizados"
         Height          =   375
         Left            =   4560
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label LblTotalCheque 
         Caption         =   "Total Cheque R$"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label LblTotalCC 
         Caption         =   "Total Cartão Crédito R$"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label LblTotalCD 
         Caption         =   "Total Cartão Débito R$"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label LblTotalDinheiro 
         Caption         =   "Total Dinheiro R$"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox TxtStatusPedido 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Height          =   405
      Left            =   7560
      TabIndex        =   22
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox TxtValorAReceber 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   20
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox TxtValorPago 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      TabIndex        =   19
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox TxtOBS 
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
      Left            =   960
      TabIndex        =   17
      Top             =   1920
      Width           =   9135
   End
   Begin VB.TextBox TxtValorTotal 
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
      Left            =   2640
      TabIndex        =   14
      Top             =   5040
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexItens 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4048
      _Version        =   393216
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
      Left            =   4200
      TabIndex        =   12
      Top             =   360
      Width           =   855
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
      Height          =   330
      Left            =   9120
      TabIndex        =   10
      Top             =   1440
      Width           =   975
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
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxtNumeroPedido 
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
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MskDataEntrega 
      Height          =   330
      Left            =   1200
      TabIndex        =   42
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSMask.MaskEdBox MskDataLimDev 
      Height          =   330
      Left            =   3720
      TabIndex        =   43
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   6480
      TabIndex        =   44
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label LblCliente 
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
      Height          =   330
      Left            =   2400
      TabIndex        =   40
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label LblNomeVendedor 
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
      Height          =   330
      Left            =   5160
      TabIndex        =   39
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label LblDataDevolucao 
      Alignment       =   2  'Center
      Caption         =   "Data Devolução"
      Height          =   375
      Left            =   5400
      TabIndex        =   36
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblStatusPedido 
      Caption         =   "Status Pedido"
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
      Left            =   8160
      TabIndex        =   21
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label LblValorAReceber 
      Caption         =   "Valor À Receber R$"
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label LblValorPago 
      Caption         =   "Valor Pago R$"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label LblValorTotal 
      Caption         =   "Valor Total Pedido R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label LblPedidoVendedor 
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Label LblDiasAtraso 
      Caption         =   "Dias Atraso"
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LblDataLimiteDev 
      Caption         =   "Data Limite Devolução"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblDataEntrega 
      Caption         =   "Data Entrega"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label LblDataEmissao 
      Caption         =   "Data Emissão"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód Cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblNumeroPedido 
      Caption         =   "Pedido nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmComEmissaoConsultaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdLimparTela_Click()

    TxtNumeroPedido.Text = ""
    TxtCodVendedor.Text = ""
    LblNomeVendedor.Caption = ""
    TxtStatusPedido.Text = ""
    TxtCodCliente.Text = ""
    LblCliente.Caption = ""
    MskDataEmissao.Mask = ""
    MskDataEmissao.Text = ""
    MskDataEmissao.Mask = "##/##/####"
    MskDataEntrega.Mask = ""
    MskDataEntrega.Text = ""
    MskDataEntrega.Mask = "##/##/####"
    MskDataLimDev.Mask = ""
    MskDataLimDev.Text = ""
    MskDataLimDev.Mask = "##/##/####"
    MskDataDev.Mask = ""
    MskDataDev.Text = ""
    MskDataDev.Mask = "##/##/####"
    TxtDiasAtraso.Text = ""
    TxtOBS.Text = ""
    Call formata_flex
    Call formata_flex2
    TxtValorAReceber.Text = ""
    TxtValorPago.Text = ""
    TxtValorTotal.Text = ""
    TxtTotalDinheiro.Text = ""
    TxtTotalCC.Text = ""
    TxtTotalCD.Text = ""
    TxtTotalCheque.Text = ""
    TxtDesconto.Text = ""
    TxtJuros.Text = ""

    TxtNumeroPedido.Enabled = True
    TxtNumeroPedido.SetFocus



End Sub

Public Sub TxtNumeroPedido_KeyPress(KeyAscii As Integer)
    Dim I As Integer
    Dim X, DH, CD, CC, CH, DESCONTO, JUROS As Double
    Dim VAREC As Double
    DH = 0
    CD = 0
    CC = 0
    CH = 0

    If KeyAscii = 13 And IsNumeric(TxtNumeroPedido.Text) <> Empty Then

        TxtNumeroPedido.Enabled = False

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1
        Set REG2 = New ADODB.Recordset
        REG2.ActiveConnection = CN1
        Set REG3 = New ADODB.Recordset
        REG3.ActiveConnection = CN1


        reg.Open ("SELECT P.NumPed,P.CodVend,F.Nome AS Vendedor,P.Status,P.CodCli,C.Nome AS Cliente,P.DataEmissao,P.DataEntrega,P.DataLimDev,P.DataDev,P.OBS,P.ValorT,P.ValorP FROM PEDIDOS AS P " & _
                  "INNER JOIN FUNCIONARIOS AS F ON P.CodVend = F.CodFunc " & _
                  "INNER JOIN CLIENTES AS C ON C.CodCli =P.CodCli " & _
                  "WHERE NumPed = " & Trim(TxtNumeroPedido.Text) & "")

        REG2.Open ("SELECT I.CodProd,P.Descricao,P.Preco,Quant,I.Status FROM ITENS AS I " & _
                   "INNER JOIN PRODUTOS AS P on I.CodProd = P.CodProd WHERE NUMPED = " & Trim(TxtNumeroPedido.Text) & "")

        REG3.Open ("SELECT P.VDinheiro,P.VCCredito,P.VCDebito,P.VCheque,P.Juros,P.Desconto,P.CodCupom,C.Tipo,C.Valor FROM PAGAMENTOS AS P " & _
                   "FULL OUTER JOIN CUPONS AS C ON P.CodCupom=C.CodCupom " & _
                   "WHERE P.NUMPED = " & Trim(TxtNumeroPedido.Text) & "")



        If reg.EOF = False Then

            TxtCodVendedor.Text = reg.Fields("CodVend")
            LblNomeVendedor.Caption = reg.Fields("Vendedor")
            TxtStatusPedido.Text = reg.Fields("Status")
            TxtCodCliente.Text = reg.Fields("CodCli")
            LblCliente.Caption = reg.Fields("Cliente")
            MskDataEmissao.Text = Format(reg.Fields("DataEmissao"), "DD/MM/YYYY")
            MskDataEntrega.Text = Format(reg.Fields("DataEntrega"), "DD/MM/YYYY")
            MskDataLimDev.Text = Format(reg.Fields("DataLimDev"), "DD/MM/YYYY")
            TxtValorPago.Text = Format(reg.Fields("ValorP"), "#,##0.00")
            TxtValorTotal.Text = Format(reg.Fields("ValorT"), "#,##0.00")



            If reg.Fields("DataDev") = "01/01/1900" Then


                MskDataDev.Mask = ""
                MskDataDev.Text = ""
                MskDataDev.Mask = "##/##/####"
                I = DateDiff("d", Format(reg.Fields("DataLimDev"), "DD/MM/YYYY"), Now)
                If I < 0 Then
                    I = 0
                End If

            Else

                MskDataDev.Text = Format(reg.Fields("DataDev"), "DD/MM/YYYY")
                I = DateDiff("d", Format(reg.Fields("DataLimDev"), "DD/MM/YYYY"), Format(reg.Fields("DataDev"), "DD/MM/YYYY"))
                If I < 0 Then
                    I = 0
                End If

            End If

            TxtDiasAtraso.Text = I
            TxtOBS.Text = reg.Fields("OBS")

            Call formata_flex

            Do Until REG2.EOF = True

                X = REG2.Fields("Quant") * REG2.Fields("preco")

                MSFlexItens.AddItem (REG2.Fields("codProd") & vbTab & _
                                     REG2.Fields("Descricao") & vbTab & _
                                     REG2.Fields("Quant") & vbTab & _
                                     Format(REG2.Fields("Preco"), "#,##0.00") & vbTab & _
                                     Format(X, "#,##0.00"))

                REG2.MoveNext

            Loop

            Call formata_flex2
            DESCONTO = 0
            Do Until REG3.EOF = True




                DH = DH + REG3.Fields("VDinheiro")
                CD = CD + REG3.Fields("VCDebito")
                CC = CC + REG3.Fields("VCCredito")
                CH = CH + REG3.Fields("VCheque")
                JUROS = JUROS + REG3("Juros")


                If REG3.Fields("Desconto") = Empty Or REG3.Fields("Desconto") = Null Then

                    DESCONTO = DESCONTO + 0

                Else


                    DESCONTO = DESCONTO + REG3.Fields("Desconto")


                End If

                If REG3.Fields("Codcupom") <> Empty Then

                    If REG3.Fields("Tipo") = "V" Then
                        'DESCONTO = DESCONTO + REG3.Fields("Valor")
                        MSFlexCupons.AddItem (REG3.Fields("CodCupom") & vbTab & _
                                              "R$ " & Format(REG3.Fields("Valor"), "#,##0.00"))

                    Else
                        'DESCONTO = DESCONTO + (REG3.Fields("Valor") * reg.Fields("ValorT"))

                        MSFlexCupons.AddItem (REG3.Fields("CodCupom") & vbTab & _
                                              Format((REG3.Fields("Valor") * 100), "#,##0.00") & " %")

                    End If

                End If


                REG3.MoveNext

            Loop

            TxtTotalDinheiro = Format(DH, "#,##0.00")
            TxtTotalCD = Format(CD, "#,##0.00")
            TxtTotalCheque = Format(CH, "#,##0.00")
            TxtTotalCC = Format(CC, "#,##0.00")
            TxtJuros.Text = Format(JUROS, "#,##0.00")
            TxtDesconto.Text = Format(DESCONTO, "#,##0.00")

            If JUROS = Empty Then
                JUROS = 0
            End If

            If DESCONTO = Empty Then
                DESCONTO = 0
            End If



            VAREC = CDbl(TxtValorTotal.Text) - CDbl(TxtValorPago.Text) + JUROS
            TxtValorAReceber = Format(VAREC, "#,##0.00")

        Else

            MsgBox "Pedido não encontrado", vbInformation, Aviso

        End If

        reg.Close
        REG2.Close
        REG3.Close

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
    MSFlexItens.Text = "Descrição"
    MSFlexItens.ColWidth(1) = 4500

    MSFlexItens.Col = 2
    MSFlexItens.Text = "Qtd"
    MSFlexItens.ColWidth(2) = 750


    MSFlexItens.Col = 3
    MSFlexItens.Text = "Valor Unit."
    MSFlexItens.ColWidth(3) = 1000

    MSFlexItens.Col = 4
    MSFlexItens.Text = "valor Total"
    MSFlexItens.ColWidth(4) = 1000


End Sub
Private Sub formata_flex2()

    MSFlexCupons.Clear
    MSFlexCupons.Cols = 2
    MSFlexCupons.Rows = 1

    MSFlexCupons.Col = 0
    MSFlexCupons.Text = "Cód."
    MSFlexCupons.ColWidth(0) = 1000

    MSFlexCupons.Col = 1
    MSFlexCupons.Text = "Valor"
    MSFlexCupons.ColWidth(1) = 1000


End Sub
