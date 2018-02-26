VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCapLancBuscarDocto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busca por Nº Documento"
   ClientHeight    =   6120
   ClientLeft      =   17355
   ClientTop       =   5145
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNumDocto 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPesquisa 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmComCapLancBuscarDocto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = 1500
    Me.Left = 1500
End Sub
Private Sub CmdBuscar_Click()

    If ValidaCampos() = True Then


        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1

        reg.Open ("SELECT P.NumDocto,P.CodForn,F.RazaoSocial,P.Valor,P.Desconto,P.Juros,P.Status FROM C_A_P AS P INNER JOIN FORNECEDORES AS F ON P.CodForn = F.CodForn WHERE NumDocTO LIKE '" & Trim(TxtNumDocto.Text) & "%'")



        Call formata_flex

        Do Until reg.EOF = True

            MSFlexPesquisa.AddItem (reg.Fields("numdocto") & vbTab & _
                                    reg.Fields("codforn") & vbTab & _
                                    reg.Fields("razaosocial") & vbTab & _
                                    Format(reg.Fields("valor") - reg.Fields("desconto") + reg.Fields("juros"), "#,##0.00") & vbTab & _
                                    reg.Fields("status"))

            reg.MoveNext

        Loop

        reg.Close
    End If

End Sub
Private Sub formata_flex()

    MSFlexPesquisa.Clear
    MSFlexPesquisa.Cols = 5
    MSFlexPesquisa.Rows = 1

    MSFlexPesquisa.Col = 0
    MSFlexPesquisa.Text = "Num.Docto"
    MSFlexPesquisa.ColWidth(0) = 1300

    MSFlexPesquisa.Col = 1
    MSFlexPesquisa.Text = "Cod.Forn."
    MSFlexPesquisa.ColWidth(1) = 700

    MSFlexPesquisa.Col = 2
    MSFlexPesquisa.Text = "Razão Social"
    MSFlexPesquisa.ColWidth(2) = 3000

    MSFlexPesquisa.Col = 3
    MSFlexPesquisa.Text = "Valor Total"
    MSFlexPesquisa.ColWidth(3) = 700

    MSFlexPesquisa.Col = 4
    MSFlexPesquisa.Text = "Status"
    MSFlexPesquisa.ColWidth(4) = 500

End Sub
Private Sub CmdLimparTela_Click()

    Call limpa_campos

End Sub
Private Sub limpa_campos()

    TxtNumDocto.Text = ""
    Call formata_flex

    TxtNumDocto.SetFocus
End Sub

Private Function ValidaCampos() As Boolean


    If TxtNumDocto.Text <> Empty Then

        ValidaCampos = True

    Else

        ValidaCampos = False

    End If


End Function

Private Sub MSFlexPesquisa_KeyPress(KeyAscii As Integer)

    Dim CODIGO As String

    If KeyAscii = 13 Then

        MSFlexPesquisa.Col = 0
        CODIGO = Trim(MSFlexPesquisa.Text)

        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set REG3 = New ADODB.Recordset
        REG3.ActiveConnection = CN1

        REG3.Open ("SELECT * FROM C_A_P WHERE NumDocTO LIKE '" & CODIGO & "%'")

        Call FrmComCapLanc.limpa_campos
        FrmComCapLanc.TxtCodForn.Text = REG3.Fields("CodForn")
        FrmComCapLanc.TxtCodForn_KeyPress (13)
        FrmComCapLanc.MskVencto.Text = Format(REG3.Fields("vencto"), "DD/MM/YYYY")
        FrmComCapLanc.TxtSeq.Text = REG3.Fields("seq")
        FrmComCapLanc.TxtSeq_KeyPress (13)

        REG3.Close


        Unload Me



    End If

End Sub

Private Sub TxtNumDocto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtNumDocto.Text <> Empty Then

        Call CmdBuscar_Click

        MSFlexPesquisa.SetFocus

    End If

End Sub






