VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComCapLanc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento Contas a Pagar"
   ClientHeight    =   5415
   ClientLeft      =   1320
   ClientTop       =   2565
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8400
   Begin MSMask.MaskEdBox MskVencto 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
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
   Begin VB.CommandButton CmdBuscarNumDocto 
      Caption         =   "Buscar por Nº Documento"
      Height          =   375
      Left            =   2400
      TabIndex        =   31
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox TxtSeq 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton CmdBuscarNome 
      Caption         =   "Buscar Por Nome"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox TxtLocalPagto 
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
      Left            =   1560
      TabIndex        =   11
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Gravar"
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox TxtOBS 
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
      TabIndex        =   12
      Top             =   3960
      Width           =   7215
   End
   Begin VB.Frame FrameValor 
      Height          =   2175
      Left            =   5280
      TabIndex        =   21
      Top             =   1440
      Width           =   2895
      Begin VB.TextBox TxtValorTotal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtJuros 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtDesconto 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LblValorTotal 
         Caption         =   "Valor Total R$"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label LblJuros 
         Caption         =   "Juros/Multa R$"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblDesconto 
         Caption         =   "Desconto R$"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.Label LblValor 
         Caption         =   "Valor R$"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox TxtStatus 
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
      Left            =   1560
      TabIndex        =   9
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox TxtNumDocto 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ComboBox CmbTipoPagto 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox TxtCodForn 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MskDataPagto 
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin MSMask.MaskEdBox MskDataLancto 
      Height          =   375
      Left            =   6840
      TabIndex        =   32
      Top             =   240
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
   Begin VB.Label LblForn 
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
      Left            =   3000
      TabIndex        =   33
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label LblSeq 
      Caption         =   "Seq."
      Height          =   255
      Left            =   3000
      TabIndex        =   30
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label LblLocalPagto 
      Caption         =   "Local Pagto"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label LblDataLancto 
      Caption         =   "Data Lançamento"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label LblOBS 
      Caption         =   "OBS"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label LblTipoPagto 
      Caption         =   "Tipo Pagamento"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LblNumDocto 
      Caption         =   "Nº Documento"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label LblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label LblDataPagto 
      Caption         =   "Data Pagto"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LblVencto 
      Caption         =   "Vencimento"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label LblCodForn 
      Caption         =   "Cod. Forn"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "FrmComCapLanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTOTAL, VALOR, DESCONTO, JUROS As Double
Attribute VALOR.VB_VarUserMemId = 1073938432
Attribute DESCONTO.VB_VarUserMemId = 1073938432
Attribute JUROS.VB_VarUserMemId = 1073938432

Private Sub CmdBuscarNome_Click()

    FrmComCapLancBuscarNome.Show

End Sub

Private Sub CmdBuscarNumDocto_Click()

    FrmComCapLancBuscarDocto.Show

End Sub
Private Sub carregar_combo_tipo()

    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1

    reg.Open ("SELECT Descricao FROM TIPOS_PAGTOS_CAP order by descricao")

    Do Until reg.EOF = True

        CmbTipoPagto.AddItem (reg.Fields("Descricao"))

        reg.MoveNext

    Loop

    reg.Close

End Sub
Public Sub limpa_campos()

    MskDataLancto.Mask = ""
    MskDataLancto.Text = ""
    MskDataLancto.Mask = "##/##/####"
    TxtCodForn.Text = ""
    LblForn.Caption = ""
    MskVencto.Mask = ""
    MskVencto.Text = ""
    MskVencto.Mask = "##/##/####"
    TxtSeq.Text = ""
    TxtSeq.Locked = False
    TxtNumDocto.Text = ""
    CmbTipoPagto.Text = ""
    TxtValor = ""
    TxtDesconto.Text = ""
    TxtJuros.Text = ""
    TxtValor.Text = ""
    TxtValorTotal = ""
    TxtStatus.Text = ""
    MskDataPagto.Mask = ""
    MskDataPagto.Text = ""
    MskDataPagto.Mask = "##/##/####"
    TxtLocalPagto.Text = ""
    TxtOBS.Text = ""


    VTOTAL = 0
    VALOR = 0
    DESCONTO = 0
    JUROS = 0

    TxtCodForn.SetFocus


End Sub
Private Sub valor_total()

    VTOTAL = VALOR - DESCONTO + JUROS

    TxtValorTotal.Text = Format(VTOTAL, "#,##0.00")


End Sub

Private Function ValidaCampos() As Boolean


    If IsNumeric(TxtCodForn.Text) <> Empty And LblForn.Caption <> Empty Then

        ValidaCampos = True

        If IsDate(MskVencto.Text) <> Empty Then

            ValidaCampos = True

            If IsNumeric(TxtSeq.Text) <> Empty Then

                ValidaCampos = True

                If TxtNumDocto.Text <> Empty Then

                    ValidaCampos = True

                    If CmbTipoPagto.Text <> Empty Then

                        ValidaCampos = True

                        If VALOR <> Empty Then

                            ValidaCampos = True

                            If VTOTAL >= 0 Then

                                ValidaCampos = True

                                If TxtStatus.Text <> Empty Then

                                    ValidaCampos = True

                                    If ((TxtStatus.Text = "A" Or TxtStatus.Text = "C") And Replace(Replace(MskDataPagto.Text, "/", ""), "_", "") = Empty) Or (TxtStatus.Text = "P" And IsDate(MskDataPagto.Text) <> Empty) Then

                                        ValidaCampos = True

                                        If IsDate(MskDataLancto.Text) <> Empty Then

                                            ValidaCampos = True

                                        Else

                                            ValidaCampos = False

                                        End If

                                    Else

                                        ValidaCampos = False

                                        MsgBox "Data do Pagto Incorreta"

                                    End If

                                Else

                                    ValidaCampos = False

                                End If

                            Else

                                ValidaCampos = False
                                MsgBox "Valor Total Menor ou Igual a Zero"

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

Private Sub CmdGravar_Click()

    Call ValidaCampos

    If ValidaCampos = True Then

        Dim PAGTO As Date
        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1


        reg.Open ("SELECT CodForn,Vencto,Seq FROM C_A_P WHERE CODFORN = " & Trim(TxtCodForn.Text) & " AND Vencto = '" & Format(MskVencto.Text, "YYYYMMDD") & "' AND SEQ = " & Trim(TxtSeq.Text) & "")


        If StrConv(TxtStatus.Text, vbUpperCase) = "A" Or StrConv(TxtStatus.Text, vbUpperCase) = "C" Then

            PAGTO = "01/01/1900"

        Else

            PAGTO = Format(MskDataPagto.Text, "DD/MM/YYYY")

        End If



        If reg.EOF = True Then

            CN1.Execute ("INSERT INTO C_A_P(CodForn,Vencto,Seq,DataLancto,NumDocto,Tipo,Valor,Desconto,Juros,Status,DataPagto,LocalPagto,OBS,Usuario,DataEmissao) " & _
                         "VALUES(" & TxtCodForn.Text & ",'" & Format(MskVencto.Text, "YYYYMMDD") & "'," & TxtSeq.Text & ",'" & Format(MskDataLancto.Text, "YYYYMMDD") & "','" & _
                         StrConv(TxtNumDocto.Text, vbUpperCase) & "','" & StrConv(CmbTipoPagto.Text, vbUpperCase) & "'," & Replace(VALOR, ",", ".") & "," & Replace(DESCONTO, ",", ".") & "," & _
                         Replace(JUROS, ",", ".") & ",'" & StrConv(TxtStatus.Text, vbUpperCase) & "','" & Format(PAGTO, "YYYYMMDD") & "','" & _
                         StrConv(TxtLocalPagto.Text, vbUpperCase) & "','" & StrConv(TxtOBS.Text, vbUpperCase) & "','','" & Format(Now, "YYYYMMDD hh:mm") & "')")

        Else

            CN1.Execute ("UPDATE C_A_P SET " & _
                         "DataLancto = '" & Format(MskDataLancto.Text, "YYYYMMDD") & "',NumDocto = '" & StrConv(TxtNumDocto.Text, vbUpperCase) & "',Tipo='" & StrConv(CmbTipoPagto.Text, vbUpperCase) & "',Valor=" & Replace(VALOR, ",", ".") & ", " & _
                         "Desconto= " & Replace(DESCONTO, ",", ".") & ",Juros=" & Replace(JUROS, ",", ".") & ",Status='" & StrConv(TxtStatus.Text, vbUpperCase) & "',DataPagto='" & Format(PAGTO, "YYYYMMDD") & "'," & _
                         "LocalPagto ='" & StrConv(TxtLocalPagto.Text, vbUpperCase) & "',OBS='" & StrConv(TxtOBS.Text, vbUpperCase) & "',Usuario= '',DataEmissao = '" & Format(Now, "YYYYMMDD hh:mm") & "' " & _
                         "WHERE CodForn= " & Trim(TxtCodForn.Text) & " AND Vencto = '" & Format(MskVencto.Text, "YYYYMMDD") & "' AND Seq=" & Trim(TxtSeq.Text) & "")

        End If

        Call limpa_campos


    Else


        MsgBox "Verifique os Campos"


    End If


End Sub

Private Sub CmdLimparTela_Click()

    Call limpa_campos

End Sub

Private Sub Form_Load()

    Call carregar_combo_tipo

End Sub

Public Sub TxtCodForn_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 And IsNumeric(TxtCodForn.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1


        reg.Open ("SELECT RazaoSocial FROM FORNECEDORES WHERE CODFORN = " & Trim(TxtCodForn.Text) & "")

        If reg.EOF = False Then

            LblForn.Caption = reg.Fields("RazaoSocial")
            MskVencto.SetFocus

        Else

            MsgBox "Fornecedor não encontrado"
            LblForn.Caption = ""

        End If

        reg.Close

    End If


End Sub

Private Sub MskVencto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsDate(MskVencto.Text) <> Empty Then

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1
        Dim SEQ As Integer

        reg.Open ("SELECT Seq FROM C_A_P WHERE VENCTO = '" & Format(MskVencto.Text, "YYYYMMDD") & "' and CODFORN = " & Trim(TxtCodForn.Text) & " order by seq desc")

        If reg.EOF = True Then

            SEQ = 1

        Else

            SEQ = reg.Fields("seq") + 1

        End If

        reg.Close

        TxtSeq.Text = SEQ
        MskDataLancto = Format(Now, "DD/MM/YYYY")
        TxtSeq.SetFocus
        TxtSeq.Locked = False
        TxtSeq.SelStart = 0
        TxtSeq.SelLength = Len(TxtSeq.Text)


    End If


End Sub
Public Sub TxtSeq_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtSeq.Text) <> Empty Then


        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set REG2 = New ADODB.Recordset
        REG2.ActiveConnection = CN1

        REG2.Open ("SELECT * FROM C_A_P WHERE CODFORN = " & Trim(TxtCodForn.Text) & " and VENCTO = '" & Format(MskVencto.Text, "YYYYMMDD") & "' " & _
                   " AND SEQ = " & Trim(TxtSeq.Text) & "")


        If REG2.EOF = False Then

            TxtCodForn_KeyPress (13)

            MskDataLancto.Enabled = True

            MskDataLancto.Mask = ""
            MskDataLancto.Text = Format(REG2.Fields("DataLancto"), "DD/MM/YYYY")
            MskDataLancto.Mask = "##/##/####"
            TxtNumDocto.Text = REG2.Fields("NumDocto")
            CmbTipoPagto.Text = REG2.Fields("Tipo")
            TxtValor.Text = Format(REG2.Fields("Valor"), "#,##0.00")
            TxtDesconto.Text = Format(REG2.Fields("Desconto"), "#,##0.00")
            TxtJuros.Text = Format(REG2.Fields("Juros"), "#,##0.00")

            VALOR = REG2.Fields("Valor")
            DESCONTO = REG2.Fields("Desconto")
            JUROS = REG2.Fields("Juros")

            Call valor_total

            TxtStatus.Text = REG2.Fields("Status")

            If TxtStatus.Text = "P" Then
                MskDataPagto.Text = Format(REG2.Fields("DataPagto"), "DD/MM/YYYY")
                TxtLocalPagto.Text = REG2.Fields("LocalPagto")
            End If

            TxtOBS.Text = REG2.Fields("OBS")

            REG2.Close

            TxtSeq.Locked = True
            TxtNumDocto.SetFocus
            MskDataLancto.Enabled = False

        Else

            TxtNumDocto.SetFocus

        End If


    End If

End Sub



Private Sub TxtNumDocto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtNumDocto.Text <> Empty Then

        CmbTipoPagto.SetFocus


    End If

End Sub

Private Sub CmbTipoPagto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And CmbTipoPagto.Text <> Empty Then

        TxtValor.SetFocus

    End If


End Sub





Private Sub TxtValor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsNumeric(TxtValor.Text) <> Empty Then

        VALOR = CDbl(TxtValor.Text)
        TxtDesconto.SetFocus
        TxtValor = Format(VALOR, "#,##0.00")
        Call valor_total

    End If


End Sub
Private Sub TxtDesconto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsNumeric(TxtDesconto.Text) <> Empty Or TxtDesconto.Text = Empty) Then

        If TxtDesconto.Text = Empty Then

            DESCONTO = 0

        Else

            DESCONTO = CDbl(TxtDesconto.Text)

        End If

        TxtJuros.SetFocus
        TxtDesconto = Format(DESCONTO, "#,##0.00")
        Call valor_total

    End If


End Sub
Private Sub TxtJuros_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsNumeric(TxtJuros.Text) <> Empty Or TxtJuros.Text = Empty) Then

        If TxtJuros.Text = Empty Then

            JUROS = 0

        Else

            JUROS = CDbl(TxtJuros.Text)

        End If

        TxtStatus.SetFocus
        TxtStatus.Text = "A"
        TxtStatus.SelStart = 0
        TxtStatus.SelLength = Len(TxtStatus.Text)
        TxtJuros = Format(JUROS, "#,##0.00")
        Call valor_total

    End If


End Sub

Private Sub TxtStatus_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If StrConv(TxtStatus.Text, vbUpperCase) = "A" Or StrConv(TxtStatus.Text, vbUpperCase) = "C" Then

            TxtOBS.SetFocus
            TxtStatus.Text = StrConv(TxtStatus.Text, vbUpperCase)
            MskDataPagto.Mask = ""
            MskDataPagto.Text = ""
            MskDataPagto.Mask = "##/##/####"
            MskDataPagto.Enabled = False



        ElseIf KeyAscii = 13 And StrConv(TxtStatus.Text, vbUpperCase) = "P" Then

            MskDataPagto.Enabled = True
            MskDataPagto = Format(Now, "DD/MM/YYYY")
            TxtLocalPagto.SetFocus
            TxtStatus.Text = StrConv(TxtStatus.Text, vbUpperCase)

        Else

            MsgBox "Digite somente A - Aberto,P - Pago ou C - Cancelado"

        End If

    End If

End Sub
Private Sub TxtLocalPagto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        TxtOBS.SetFocus

    End If


End Sub
Private Sub MskDataPagto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And IsDate(MskDataPagto.Text) <> Empty Then

        TxtLocalPagto.SetFocus

    End If


End Sub


Private Sub TxtOBS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        CmdGravar.SetFocus

    End If


End Sub


