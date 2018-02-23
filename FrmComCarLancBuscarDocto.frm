VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCarLancBuscarDocto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar por Nº Documento"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6750
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox TxtNumDocto 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPesquisa 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmComCarLancBuscarDocto"
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

            reg.Open ("SELECT R.NumDocto,R.CodCli,C.Nome,R.Valor,R.Status FROM C_A_R AS R INNER JOIN CLIENTES AS C ON C.CodCli = R.CodCli WHERE NumDocTO LIKE '" & Trim(TxtNumDocto.Text) & "%'")



        Call formata_flex

        Do Until reg.EOF = True

            MSFlexPesquisa.AddItem (reg.Fields("numdocto") & vbTab & _
                                    reg.Fields("codcli") & vbTab & _
                                    reg.Fields("nome") & vbTab & _
                                    Format(reg.Fields("valor"), "#,##0.00") & vbTab & _
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
    MSFlexPesquisa.Text = "Cod.Cli."
    MSFlexPesquisa.ColWidth(1) = 700
    
    MSFlexPesquisa.Col = 2
    MSFlexPesquisa.Text = "Nome"
    MSFlexPesquisa.ColWidth(2) = 3000
    
    MSFlexPesquisa.Col = 3
    MSFlexPesquisa.Text = "Valor"
    MSFlexPesquisa.ColWidth(3) = 700
    
    MSFlexPesquisa.Col = 4
    MSFlexPesquisa.Text = "Status"
    MSFlexPesquisa.ColWidth(4) = 500

End Sub
Private Sub CmdLimparTela_Click()

    Call limpa_campos
    
End Sub
Private Sub limpa_campos()

    CBoxAleatorio.Value = vbUnchecked
    TxtNumDocto.Text = ""
    Call formata_flex

    TxtDescricao.SetFocus
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
        Set REG2 = New ADODB.Recordset
        REG2.ActiveConnection = CN1

        REG2.Open ("SELECT * FROM C_A_R WHERE NumDocTO LIKE '" & CODIGO & "%'")

            Call FrmComCarLanc.limpa_campos
            FrmComCarLanc.TxtCodCliente.Text = REG2.Fields("CodCli")
            FrmComCarLanc.TxtCodCliente_KeyPress (13)
            FrmComCarLanc.MskVencto.Text = Format(REG2.Fields("vencto"), "DD/MM/YYYY")
            FrmComCarLanc.TxtSeq.Text = REG2.Fields("seq")
            FrmComCarLanc.MskDataLancto.Text = Format(REG2.Fields("DataLancto"), "DD/MM/YYYY")
            FrmComCarLanc.TxtNumeroPedido.Text = REG2.Fields("numped")
            FrmComCarLanc.TxtNumeroDocto.Text = REG2.Fields("NumDocto")
            FrmComCarLanc.CmbTipoDocto.Text = REG2.Fields("Tipo")
            FrmComCarLanc.TxtValorTotal.Text = Format(REG2.Fields("Valor"), "#,##0.00")
            FrmComCarLanc.TxtStatus.Text = REG2.Fields("Status")

            If FrmComCarLanc.TxtStatus.Text = "P" Then
                FrmComCarLanc.MskDataPagto.Text = Format(REG2.Fields("DataPagto"), "DD/MM/YYYY")
            End If

            FrmComCarLanc.TxtOBS.Text = REG2.Fields("OBS")


        REG2.Close

        Unload Me

    End If

End Sub

Private Sub TxtNumDocto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtNumDocto.Text <> Empty Then

        Call CmdBuscar_Click

        MSFlexPesquisa.SetFocus

    End If

End Sub




