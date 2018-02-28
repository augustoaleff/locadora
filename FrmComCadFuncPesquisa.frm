VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmComCadFuncPesquisa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca por Nome"
   ClientHeight    =   7440
   ClientLeft      =   4440
   ClientTop       =   735
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   6795
   Begin VB.CheckBox CBoxIncluirDesligados 
      Caption         =   "Incluir Desligados?"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame FrameBuscarPor 
      Caption         =   "Buscar Por"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6375
      Begin VB.OptionButton OptCEP 
         Caption         =   "CEP"
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptRG 
         Caption         =   "RG"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptCPF 
         Caption         =   "CPF"
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton OptNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox TxtNome 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox CBoxAleatorio 
      Caption         =   "Nome Aleatório"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CmdLimparTela 
      Caption         =   "Limpar Tela"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPesquisa 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmComCadFuncPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = 1500
    Me.Left = 1500
End Sub


Private Sub CmdBuscar_Click()

    If Valida_Campos() = True Then


        Set CN1 = New ADODB.Connection
        CN1.Open STR_DSN
        Set reg = New ADODB.Recordset
        reg.ActiveConnection = CN1
        
        Dim DESLIG As String
        
        If CBoxIncluirDesligados.Value = vbChecked Then
        
            DESLIG = ""
        
        
        Else
        
            DESLIG = " AND DATADEM = '19000101'"
        
        End If


        If OptNome.Value = True Then

            If CBoxAleatorio.Value = vbChecked Then

                reg.Open ("SELECT * FROM FUNCIONARIOS WHERE NOME LIKE '%" & Trim(TxtNome.Text) & "%'" & DESLIG & "")

            Else

                reg.Open ("SELECT * FROM FUNCIONARIOS WHERE NOME LIKE '" & Trim(TxtNome.Text) & "%'" & DESLIG & "")

            End If

        End If

        If OptCPF.Value = True Then
            reg.Open ("SELECT * FROM FUNCIONARIOS WHERE CPF LIKE '" & Replace(Replace(Replace(Replace(Replace(TxtNome.Text, ".", ""), " ", ""), "/", ""), "\", ""), "-", "") & "%'" & DESLIG & "")
        End If

        If OptRG.Value = True Then
            reg.Open ("SELECT * FROM FUNCIONARIOS WHERE RG LIKE '" & Replace(Replace(Replace(Replace(Replace(TxtNome.Text, ".", ""), " ", ""), "/", ""), "\", ""), "-", "") & "%'" & DESLIG & "")
        End If

        If OptCEP.Value = True Then
            reg.Open ("SELECT * FROM FUNCIONARIOS WHERE CEP LIKE '" & Replace(Replace(Replace(Replace(Replace(TxtNome.Text, ".", ""), " ", ""), "/", ""), "\", ""), "-", "") & "%'" & DESLIG & "")
        End If

        Call formata_flex
        

        Do Until reg.EOF = True

        If reg.Fields("DataDem") = "01/01/1900" Then
        

            MSFlexPesquisa.AddItem (reg.Fields("codfunc") & vbTab & _
                                    reg.Fields("nome") & vbTab & _
                                    vbTab)

            reg.MoveNext
            
        Else
        
            MSFlexPesquisa.AddItem (reg.Fields("codfunc") & vbTab & _
                                    reg.Fields("nome") & vbTab & _
                                    "SIM")

            reg.MoveNext
            
        End If

        Loop


        reg.Close
    End If

End Sub
Private Sub formata_flex()

    MSFlexPesquisa.Clear
    MSFlexPesquisa.Cols = 3
    MSFlexPesquisa.Rows = 1

    MSFlexPesquisa.Col = 0
    MSFlexPesquisa.Text = "Cód."
    MSFlexPesquisa.ColWidth(0) = 700

    MSFlexPesquisa.Col = 1
    MSFlexPesquisa.Text = "Nome"
    MSFlexPesquisa.ColWidth(1) = 4000
    
    MSFlexPesquisa.Col = 2
    MSFlexPesquisa.Text = "Desligado?"
    MSFlexPesquisa.ColWidth(2) = 1000

End Sub
Private Sub CmdLimparTela_Click()
    Call limpa_campos
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

        FrmComCadFunc.TxtCodigo.Text = CODIGO
        FrmComCadFunc.TxtCodigo_KeyPress (13)
        Unload Me

    End If

End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtNome.Text <> Empty Then

        Call CmdBuscar_Click

        MSFlexPesquisa.SetFocus

    End If

End Sub


