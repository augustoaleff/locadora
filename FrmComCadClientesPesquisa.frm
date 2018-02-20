VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCadClientesPesquisa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca por Nome"
   ClientHeight    =   7275
   ClientLeft      =   1635
   ClientTop       =   2865
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7020
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CheckBox CBoxAleatorio 
      Caption         =   "Nome Aleatório"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPesquisa 
      Height          =   4455
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox TxtNome 
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
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Frame FrameBuscarPor 
      Caption         =   "Buscar Por"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6375
      Begin VB.OptionButton OptCEP 
         Caption         =   "CEP"
         Height          =   255
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptRG 
         Caption         =   "RG/IE"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptCPF 
         Caption         =   "CPF/CNPJ"
         Height          =   195
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmComCadClientesPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBuscar_Click()

    If Valida_Campos() = True Then


        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1


        If OptNome.Value = True Then

            If CBoxAleatorio.Value = vbChecked Then

                reg.Open ("SELECT CODCLI,NOME FROM CLIENTES WHERE NOME LIKE '%" & Trim(TxtNome.Text) & "%'")

            Else

                reg.Open ("SELECT CODCLI,NOME FROM CLIENTES WHERE NOME LIKE '" & Trim(TxtNome.Text) & "%'")

            End If

        End If

        If OptCPF.Value = True Then
            reg.Open ("SELECT CODCLI,NOME FROM CLIENTES WHERE CPF LIKE '" & Replace(Replace(Replace(Replace(Replace(TxtNome.Text, ".", ""), " ", ""), "/", ""), "\", ""), "-", "") & "%'")
        End If

        If OptRG.Value = True Then
            reg.Open ("SELECT CODCLI,NOME FROM CLIENTES WHERE IE LIKE '" & Replace(Replace(Replace(Replace(Replace(TxtNome.Text, ".", ""), " ", ""), "/", ""), "\", ""), "-", "") & "%'")
        End If

        If OptCEP.Value = True Then
            reg.Open ("SELECT CODCLI,NOME FROM CLIENTES WHERE CEP LIKE '" & Replace(Replace(Replace(Replace(Replace(TxtNome.Text, ".", ""), " ", ""), "/", ""), "\", ""), "-", "") & "%'")
        End If

        Call formata_flex

        Do Until reg.EOF = True

            MSFlexPesquisa.AddItem (reg.Fields("codcli") & vbTab & _
                                    reg.Fields("nome"))

            reg.MoveNext

        Loop


        reg.Close
    End If

End Sub
Private Sub formata_flex()

    MSFlexPesquisa.Clear
    MSFlexPesquisa.Cols = 2
    MSFlexPesquisa.Rows = 1

    MSFlexPesquisa.Col = 0
    MSFlexPesquisa.Text = "Cód."
    MSFlexPesquisa.ColWidth(0) = 700

    MSFlexPesquisa.Col = 1
    MSFlexPesquisa.Text = "Nome"
    MSFlexPesquisa.ColWidth(1) = 5400

End Sub
Private Sub CmdLimparTela_Click()
    Call limpa_campos
End Sub

Private Sub Form_Load()
    Me.Top = 1500
    Me.Left = 1500

End Sub
Private Sub limpa_campos()
    OptNome.Value = True
    CBoxAleatorio.Value = vbUnchecked
    TxtNome.Text = ""
    Call formata_flex


    TxtNome.SetFocus
End Sub

Private Function Valida_Campos() As Boolean

    If OptCPF.Value = True Or OptRG.Value = True Or OptCEP = True Then

        If IsNumeric(Replace(Replace(Replace(Replace(Replace(TxtNome.Text, ".", ""), " ", ""), "/", ""), "\", ""), "-", "")) <> Empty Then
            Valida_Campos = True
        Else
            Valida_Campos = False
            MsgBox "Digite somente números", vbExclamation, "Aviso"
        End If

    Else

        If TxtNome.Text <> Empty Then
            Valida_Campos = True
        Else
            Valida_Campos = False

        End If

    End If


End Function

Private Sub MSFlexPesquisa_KeyPress(KeyAscii As Integer)

    Dim CODIGO As Long

    If KeyAscii = 13 Then

        MSFlexPesquisa.Col = 0
        CODIGO = Trim(MSFlexPesquisa.Text)

        FrmComCadClientes.TxtCodigo.Text = CODIGO
        FrmComCadClientes.TxtCodigo_KeyPress (13)
        Unload Me

    End If

End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtNome.Text <> Empty Then

        Call CmdBuscar_Click

        MSFlexPesquisa.SetFocus

    End If

End Sub
