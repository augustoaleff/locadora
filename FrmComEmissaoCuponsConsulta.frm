VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmComEmissaoCuponsConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cupons"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   8595
   Begin MSMask.MaskEdBox MskPeriodoDe 
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
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
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton CmdImprimirRelat 
      Caption         =   "Imprimir Relatório"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   6360
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexCupons 
      Height          =   4095
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.TextBox TxtCodCupom 
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame FramePesquisar 
      Caption         =   "Pesquisar"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton OptNaoUtilizados 
         Caption         =   "Não Utilizados"
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptUtilizados 
         Caption         =   "Utilizados"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin MSMask.MaskEdBox MskPeriodoAte 
      Height          =   330
      Left            =   2880
      TabIndex        =   2
      Top             =   1320
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
   Begin VB.Label LblA 
      Caption         =   "à"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label LblValorTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   6360
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label LblValor 
      Caption         =   "Valor Total R$"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label LblPeriodo 
      Caption         =   "Período"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label LblCodCupom 
      Caption         =   "Cód Cupom"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "FrmComEmissaoCuponsConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdConsultar_Click()

    Dim DE, ATE As Date
    Set CN1 = New ADODB.Connection
    CN1.Open STR_DSN
    Set reg = New ADODB.Recordset
    reg.ActiveConnection = CN1


    DE = Format("01/01/1900", "DD/MM/YYYY")
    ATE = Format("31/12/2199", "DD/MM/YYYY")

    If Replace(Replace(MskPeriodoDe.Text, "/", ""), "_", "") = Empty Then

        MskPeriodoDe.Text = Format(DE, "DD/MM/YYYY")

    End If

    If Replace(Replace(MskPeriodoAte.Text, "/", ""), "_", "") = Empty Then

        MskPeriodoAte.Text = Format(ATE, "DD/MM/YYYY")

    End If

    If TxtCodCupom.Text <> Empty Then



        If OptUtilizados.Value = vbTrue Then

            reg.Open ("SELECT * FROM CUPONS WHERE CODCUPOM like '%" & Trim(StrConv(TxtCodCupom.Text, vbUpperCase)) & "%' " & _
                      "and DataEmissao between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "' and " & _
                      "Status = 'UTILIZADO'")
        Else

            If OptNaoUtilizados.Value = vbTrue Then

                reg.Open ("SELECT * FROM CUPONS WHERE CODCUPOM like '%" & Trim(StrConv(TxtCodCupom.Text, vbUpperCase)) & "%' " & _
                          "and DataEmissao between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "' and " & _
                          "Status = 'NAOUTILIZADO'")

            Else

                reg.Open ("SELECT * FROM CUPONS WHERE CODCUPOM like '%" & Trim(StrConv(TxtCodCupom.Text, vbUpperCase)) & "%' " & _
                          "and DataEmissao between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "'")


            End If

        End If

    Else

        If OptUtilizados.Value = vbTrue Then

            reg.Open ("SELECT * FROM CUPONS " & _
                      "WHERE DataEmissao between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "' and " & _
                      "Status = 'UTILIZADO'")
        Else

            If OptNaoUtilizados.Value = vbTrue Then

                reg.Open ("SELECT * FROM CUPONS " & _
                          "WHERE DataEmissao between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "' and " & _
                          "Status = 'NAOUTILIZADO'")

            Else

                reg.Open ("SELECT * FROM CUPONS " & _
                          "WHERE DataEmissao between '" & Format(MskPeriodoDe.Text, "YYYYMMDD") & "' and '" & Format(MskPeriodoAte.Text, "YYYYMMDD") & "'")

            End If

        End If



    End If

    Call formata_flex

    Do Until reg.EOF = True

        If reg.Fields("Tipo") = "V" Then

            MSFlexCupons.AddItem (reg.Fields("CodCupom") & vbTab & _
                                  "R$ " & Format(reg.Fields("Valor"), "#,##0.00") & vbTab & _
                                  reg.Fields("Descricao") & vbTab & _
                                  Format(reg.Fields("ValidadeDe"), "DD/MM/YYYY") & vbTab & _
                                  Format(reg.Fields("ValidadeAte"), "DD/MM/YYYY") & vbTab & _
                                  Format(reg.Fields("DataEmissao"), "DD/MM/YYYY") & vbTab & _
                                  reg.Fields("Status"))

        Else

            MSFlexCupons.AddItem (reg.Fields("CodCupom") & vbTab & _
                                  Format(reg.Fields("Valor") * 100, "#,##0.00") & " %" & vbTab & _
                                  reg.Fields("Descricao") & vbTab & _
                                  Format(reg.Fields("ValidadeDe"), "DD/MM/YYYY") & vbTab & _
                                  Format(reg.Fields("ValidadeAte"), "DD/MM/YYYY") & vbTab & _
                                  Format(reg.Fields("DataEmissao"), "DD/MM/YYYY") & vbTab & _
                                  reg.Fields("Status"))

        End If

        reg.MoveNext

    Loop

    reg.Close

End Sub

Private Sub TxtCodCupom_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        MskPeriodoDe.SetFocus

    End If

End Sub
Private Sub MskPeriodoDe_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskPeriodoDe.Text) Or Replace(Replace(MskPeriodoDe.Text, "/", ""), "_", "") = Empty) Then

        MskPeriodoAte.SetFocus

    End If

End Sub
Private Sub MskPeriodoAte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And (IsDate(MskPeriodoAte.Text) Or Replace(Replace(MskPeriodoAte.Text, "/", ""), "_", "") = Empty) Then

        CmdConsultar.SetFocus

    End If

End Sub

Private Sub formata_flex()

    MSFlexCupons.Clear
    MSFlexCupons.Cols = 7
    MSFlexCupons.Rows = 1

    MSFlexCupons.Col = 0
    MSFlexCupons.Text = "Cód. Cupom"
    MSFlexCupons.ColWidth(0) = 1000

    MSFlexCupons.Col = 1
    MSFlexCupons.Text = "Valor"
    MSFlexCupons.ColWidth(1) = 1000

    MSFlexCupons.Col = 2
    MSFlexCupons.Text = "Descricao"
    MSFlexCupons.ColWidth(2) = 1500


    MSFlexCupons.Col = 3
    MSFlexCupons.Text = "Valido De"
    MSFlexCupons.ColWidth(3) = 1000

    MSFlexCupons.Col = 4
    MSFlexCupons.Text = "Valido Ate"
    MSFlexCupons.ColWidth(4) = 1000

    MSFlexCupons.Col = 5
    MSFlexCupons.Text = "Data Emisssao"
    MSFlexCupons.ColWidth(5) = 1000

    MSFlexCupons.Col = 6
    MSFlexCupons.Text = "Status"
    MSFlexCupons.ColWidth(6) = 1500


End Sub


