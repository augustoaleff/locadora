VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCarConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Contas a Receber"
   ClientHeight    =   7500
   ClientLeft      =   10125
   ClientTop       =   1815
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11655
   Begin VB.ComboBox CmbTipoPagto 
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
      ItemData        =   "FrmComCarConsulta.frx":0000
      Left            =   7440
      List            =   "FrmComCarConsulta.frx":0002
      TabIndex        =   25
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   615
      Left            =   3960
      TabIndex        =   20
      Top             =   1680
      Width           =   3375
      Begin VB.OptionButton OptPago 
         Caption         =   "Pago"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptAberto 
         Caption         =   "Aberto"
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   1920
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
      Height          =   360
      Left            =   7320
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame FrameConsultarPor 
      Caption         =   "Consultar Por"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      Begin VB.OptionButton OptVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptDataPagamento 
         Caption         =   "Data Pagamento"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OptDataLancamento 
         Caption         =   "Data Lançamento"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox TxtCodCliente 
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
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Relatório"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   9360
      TabIndex        =   3
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton CmdConsultarNome 
      Caption         =   "Buscar por Nome"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexResultado 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
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
   Begin VB.Label LblValor 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9480
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label LblTipoPagto 
      Caption         =   "Tipo Pagto"
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label LblNumeroPedido 
      Caption         =   "Pedido nº"
      Height          =   255
      Left            =   7320
      TabIndex        =   19
      Top             =   960
      Width           =   855
   End
   Begin VB.Label LblCodCliente 
      Caption         =   "Cód Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LblPeriodoA 
      Caption         =   "à"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label LblCliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2400
      TabIndex        =   15
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label LblValorTotal 
      Caption         =   "Valor Total R$"
      Height          =   255
      Left            =   9480
      TabIndex        =   14
      Top             =   885
      Width           =   1095
   End
End
Attribute VB_Name = "FrmComCarConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIPO As String

Private Sub limpa_campos()

    OptVencimento.Value = True
    TxtCodCliente.Text = ""
    LblCliente.Caption = ""
    MskPeriodoDe.Mask = ""
    MskPeriodoDe.Text = ""
    MskPeriodoDe.Mask = "##/##/####"
    MskPeriodoAte.Mask = ""
    MskPeriodoAte.Text = ""
    MskPeriodoAte.Mask = "##/##/####"
    OptTodos.Value = True
    CmbTipoPagto.Text = ""
    TxtNumeroPedido.Text = ""
    LblValor.Caption = ""
    
    Call formata_flex

    TxtCodCliente.SetFocus


End Sub

Private Sub carregar_combo_tipo()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT Descricao FROM TIPOS_PAGTOS order by descricao")

    Do Until reg.EOF = True

        CmbTipoPagto.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

End Sub
Private Sub tipo_pagto()

    If CmbTipoPagto.Text = "Dinheiro" Then

        TIPO = " AND TIPO = 'DH'"

    ElseIf CmbTipoPagto.Text = "Cartão Crédito" Then

        TIPO = " AND TIPO = 'CC'"

    ElseIf CmbTipoPagto.Text = "Cartão Débito" Then

        TIPO = " AND TIPO = 'CD'"

    ElseIf CmbTipoPagto.Text = "Cheque" Then

        TIPO = " AND TIPO = 'CHQ'"

    ElseIf CmbTipoPagto.Text = "Dinheiro e Cheque" Then

        TIPO = " AND (TIPO = 'DH' OR TIPO = 'CHQ')"

    ElseIf CmbTipoPagto.Text = "Dinheiro e CC" Then

        TIPO = " AND (TIPO = 'DH' OR TIPO = 'CC')"

    ElseIf CmbTipoPagto.Text = "Dinheiro e CD" Then

        TIPO = " AND (TIPO = 'DH' OR TIPO = 'CD')"

    ElseIf CmbTipoPagto.Text = "CD e CC" Then

        TIPO = " AND (TIPO = 'CD' OR TIPO = 'CC')"

    ElseIf CmbTipoPagto.Text = "CD e Cheque" Then

        TIPO = " AND (TIPO = 'CD' OR TIPO = 'CHQ')"

    ElseIf CmbTipoPagto.Text = "CC e Cheque" Then

        TIPO = " AND (TIPO = 'CC' OR TIPO = 'CHQ')"

    ElseIf CmbTipoPagto.Text = "Dinheiro,CC e CD" Then

        TIPO = " AND (TIPO = 'DH' OR TIPO = 'CC' OR TIPO = 'CD')"

    ElseIf CmbTipoPagto.Text = "Dinheiro, CC e Cheque" Then

        TIPO = " AND (TIPO = 'DH' OR TIPO = 'CC' OR TIPO = 'CHQ')"

    ElseIf CmbTipoPagto.Text = "Dinheiro, CD e Cheque" Then

        TIPO = " AND (TIPO = 'DH' OR TIPO = 'CD' OR TIPO = 'CHQ')"

    ElseIf CmbTipoPagto.Text = "CD,CC e Cheque" Then

        TIPO = " AND (TIPO = 'CD' OR TIPO = 'CC' OR TIPO = 'CHQ')"

    ElseIf CmbTipoPagto.Text = "*Todos" Then

        TIPO = ""

    Else

        TIPO = ""

    End If


End Sub

Private Sub CmdConsultar_Click()

    Dim PERIODODE, PERIODOATE As Date
    Dim VTOTAL As Double

    Call tipo_pagto

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1




    If Replace(Replace(MskPeriodoDe.Text, "_", ""), "/", "") = Empty Then

        PERIODODE = "01/01/1900"

    Else

        If IsDate(MskPeriodoDe.Text) = True Then
            PERIODODE = Format(MskPeriodoDe.Text, "DD/MM/YYYY")
        Else
            MsgBox "Digite a Data Inicial Correta"
        End If

    End If


    If Replace(Replace(MskPeriodoAte.Text, "_", ""), "/", "") = Empty Then

        PERIODOATE = "31/12/2199"

    Else

        If IsDate(MskPeriodoAte.Text) = True Then
            PERIODOATE = Format(MskPeriodoAte.Text, "DD/MM/YYYY")
        Else
            MsgBox "Digite a Data Final Correta"
        End If




    End If

    If OptVencimento.Value = True Then

        If TxtCodCliente.Text <> Empty Then

            If OptAberto.Value = True Then




                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY VENCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A'" & TIPO & " ORDER BY VENCTO")

                End If


            ElseIf OptPago.Value = True Then

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY VENCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P'" & TIPO & " ORDER BY VENCTO")
                End If


            Else

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY VENCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & TIPO & " ORDER BY VENCTO")

                End If

            End If



        Else

            If OptAberto.Value = True Then


                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY VENCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A'" & TIPO & " ORDER BY VENCTO")


                End If


            ElseIf OptPago.Value = True Then

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY VENCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P'" & TIPO & " ORDER BY VENCTO")
                End If


            Else

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY VENCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE VENCTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & TIPO & " ORDER BY VENCTO")

                End If

            End If


        End If

    ElseIf OptDataPagamento.Value = True Then

        If TxtCodCliente.Text <> Empty Then

            If OptAberto.Value = True Then


                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATAPAGTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A'" & TIPO & " ORDER BY DATAPAGTO")

                End If


            ElseIf OptPago.Value = True Then

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATAPAGTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P'" & TIPO & " ORDER BY DATAPAGTO")
                End If


            Else

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATAPAGTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & TIPO & " ORDER BY DATAPAGTO")

                End If

            End If



        Else

            If OptAberto.Value = True Then


                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATAPAGTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A'" & TIPO & " ORDER BY DATAPAGTO")

                End If


            ElseIf OptPago.Value = True Then

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATAPAGTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P'" & TIPO & " ORDER BY DATAPAGTO")
                End If


            Else

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATAPAGTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAPAGTO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & TIPO & " ORDER BY DATAPAGTO")

                End If

            End If


        End If

    Else

        If TxtCodCliente.Text <> Empty Then

            If OptAberto.Value = True Then


                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATALANCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A'" & TIPO & " ORDER BY DATALANCTO")

                End If


            ElseIf OptPago.Value = True Then

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATALANCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P'" & TIPO & " ORDER BY DATALANCTO")
                End If


            Else

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATALANCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE R.CODCLI= " & Trim(TxtCodCliente.Text) & " and " & _
                              "DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & TIPO & " ORDER BY DATALANCTO")

                End If

            End If



        Else

            If OptAberto.Value = True Then


                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATALANCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'A'" & TIPO & " ORDER BY DATALANCTO")

                End If


            ElseIf OptPago.Value = True Then

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P' AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATALANCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND STATUS = 'P'" & TIPO & " ORDER BY DATALANCTO")
                End If


            Else

                If TxtNumeroPedido.Text <> Empty Then



                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & _
                              "AND NUMPED=" & Trim(TxtNumeroPedido.Text) & "" & TIPO & " ORDER BY DATALANCTO")


                Else


                    reg.Open ("SELECT R.CodCli,C.Nome,R.Vencto,R.DataLancto,R.NumPed,R.NumDocto,R.Tipo,R.Status,R.DataPagto,R.OBS,R.Valor FROM C_A_R AS R " & _
                              "FULL OUTER JOIN CLIENTES AS C ON C.CodCli = R.CodCli " & _
                              "WHERE DATAEMISSAO BETWEEN '" & Format(PERIODODE, "YYYYMMDD") & "' AND '" & Format(PERIODOATE, "YYYYMMDD") & "'" & TIPO & " ORDER BY DATALANCTO")

                End If

            End If

        End If



    End If





    Call formata_flex

    Do Until reg.EOF = True


        If reg.Fields("Status") = "P" Then


            MSFlexResultado.AddItem (reg.Fields("Vencto") & vbTab & _
                                     reg.Fields("Codcli") & vbTab & _
                                     reg.Fields("Nome") & vbTab & _
                                     reg.Fields("NumPed") & vbTab & _
                                     reg.Fields("NumDocto") & vbTab & _
                                     reg.Fields("Tipo") & vbTab & _
                                     Format(reg.Fields("Valor"), "#,##0.00") & vbTab & _
                                     reg.Fields("Status") & vbTab & _
                                     Format(reg.Fields("DataPagto"), "DD/MM/YYYY") & vbTab & _
                                     reg.Fields("Obs"))
        Else

            MSFlexResultado.AddItem (reg.Fields("Vencto") & vbTab & _
                                     reg.Fields("Codcli") & vbTab & _
                                     reg.Fields("Nome") & vbTab & _
                                     reg.Fields("NumPed") & vbTab & _
                                     reg.Fields("NumDocto") & vbTab & _
                                     reg.Fields("Tipo") & vbTab & _
                                     Format(reg.Fields("Valor"), "#,##0.00") & vbTab & _
                                     reg.Fields("Status") & vbTab & _
                                     vbTab & _
                                     reg.Fields("Obs"))

        End If


        reg.MoveNext

    Loop

    reg.Close


    For contador = 1 To MSFlexResultado.Rows - 1

        VTOTAL = VTOTAL + CDbl(MSFlexResultado.TextMatrix(contador, 6))

    Next

    LblValor.Caption = Format(VTOTAL, "#,##0.00")

End Sub

Private Sub CmdConsultarNome_Click()
    FrmComCarConsultaBuscarNome.Show
End Sub

Private Sub CmdImprimir_Click()


' Dim VETOR As Variant
'Dim Tipo As String


'If TxtTipoPagto <> Empty Then

'VETOR = Split(TxtTipoPagto.Text, ",")

'Tipo = "AND TIPO = "


'If VETOR = 0 Then

'Tipo = Tipo & VETOR1

'Else



'For Each VETOR1 In VETOR

'Tipo = Tipo + " OR TIPO = " & "'" & VETOR1 & "'"

'Next


'End If


'Else

'Tipo = ""

'End If

'MsgBox Tipo


End Sub

Private Sub CmdLimparTela_Click()
    Call limpa_campos
End Sub


Private Sub Form_Load()

    Call carregar_combo_tipo

End Sub

Public Sub TxtCodCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtCodCliente.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT Nome FROM CLIENTES WHERE CODCLI =  " & Trim(TxtCodCliente.Text) & "")

        If reg.EOF = False Then

            LblCliente.Caption = reg.Fields("Nome")

            MskPeriodoDe.SetFocus

        Else

            MsgBox "Cliente não encontrado"

        End If


        reg.Close

    ElseIf KeyAscii = 13 And TxtCodCliente.Text = Empty Then

        LblCliente.Caption = ""
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


        CmdConsultar.SetFocus


    End If


End Sub

Private Sub formata_flex()

    MSFlexResultado.Clear
    MSFlexResultado.Cols = 10
    MSFlexResultado.Rows = 1

    MSFlexResultado.Col = 0
    MSFlexResultado.Text = "Vencto"
    MSFlexResultado.ColWidth(0) = 1000

    MSFlexResultado.Col = 1
    MSFlexResultado.Text = "Cod. Cliente"
    MSFlexResultado.ColWidth(1) = 1000

    MSFlexResultado.Col = 2
    MSFlexResultado.Text = "Cliente"
    MSFlexResultado.ColWidth(2) = 3000

    MSFlexResultado.Col = 3
    MSFlexResultado.Text = "Num. Pedido"
    MSFlexResultado.ColWidth(3) = 1000

    MSFlexResultado.Col = 4
    MSFlexResultado.Text = "Num. Docto"
    MSFlexResultado.ColWidth(4) = 1500

    MSFlexResultado.Col = 5
    MSFlexResultado.Text = "Tipo"
    MSFlexResultado.ColWidth(5) = 500

    MSFlexResultado.Col = 6
    MSFlexResultado.Text = "Valor"
    MSFlexResultado.ColWidth(6) = 1000

    MSFlexResultado.Col = 7
    MSFlexResultado.Text = "Status"
    MSFlexResultado.ColWidth(7) = 500

    MSFlexResultado.Col = 8
    MSFlexResultado.Text = "Data. Pagto"
    MSFlexResultado.ColWidth(8) = 1000

    MSFlexResultado.Col = 9
    MSFlexResultado.Text = "OBS"
    MSFlexResultado.ColWidth(9) = 3000


End Sub
