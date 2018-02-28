VERSION 5.00
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000A&
   Caption         =   "Sistema Locadora"
   ClientHeight    =   9075
   ClientLeft      =   4890
   ClientTop       =   3450
   ClientWidth     =   17190
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_cadastro_geral 
      Caption         =   "Cadastro"
      Begin VB.Menu mnu_cad_clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnu_cad_produtos 
         Caption         =   "Produtos"
         Begin VB.Menu mnu_cad_produtos_novo 
            Caption         =   "Novo/Consulta"
         End
         Begin VB.Menu mnu_cad_produtos_entradas 
            Caption         =   "Entrada"
         End
      End
      Begin VB.Menu mnu_cad_func 
         Caption         =   "Funcionários"
      End
      Begin VB.Menu mnu_cad_forn 
         Caption         =   "Fornecedores"
      End
   End
   Begin VB.Menu mnu_emissoes 
      Caption         =   "Emissões"
      Begin VB.Menu mnu_emissoes_aluguel 
         Caption         =   "Aluguel de Filme"
      End
      Begin VB.Menu mnu_emissoes_devolucao 
         Caption         =   "Devolução de Filme"
      End
      Begin VB.Menu mnu_emissoes_consulta 
         Caption         =   "Consulta Pedido"
      End
      Begin VB.Menu mnu_emissoes_cupons 
         Caption         =   "Cupons"
         Begin VB.Menu mnu_emissoes_cupons_emissao 
            Caption         =   "Emissão"
         End
         Begin VB.Menu mnu_emissoes_cupons_consulta 
            Caption         =   "Consulta"
         End
      End
   End
   Begin VB.Menu mnu_cap 
      Caption         =   "Contas a Pagar"
      Begin VB.Menu mnu_cap_lanc 
         Caption         =   "Lançamento"
      End
      Begin VB.Menu mnu_cap_consulta 
         Caption         =   "Consulta"
      End
   End
   Begin VB.Menu mnu_car 
      Caption         =   "Contas a Receber"
      Begin VB.Menu mnu_car_lanc 
         Caption         =   "Lançamento"
      End
      Begin VB.Menu mnu_car_consulta 
         Caption         =   "Consulta"
      End
   End
   Begin VB.Menu mnu_relat 
      Caption         =   "Relatórios"
      Begin VB.Menu mnu_relat_ped 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu mnu_relat_cli 
         Caption         =   "Cadastro"
      End
      Begin VB.Menu mnu_relat_prod 
         Caption         =   "Produtos"
      End
      Begin VB.Menu mnu_relat_cap 
         Caption         =   "Contas A Pagar"
      End
      Begin VB.Menu mnu_relat_car 
         Caption         =   "Contas a Receber"
      End
   End
   Begin VB.Menu mnu_Sair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
'STR_DSN = "Driver={SQL Server Native Client 10.0};Server=localhost,1433;Database=Locadora;Uid=sa;Pwd=ELetro66573759000171;Connection Timeout=30;"
    STR_DSN = "Driver={SQL Server Native Client 10.0};Server=localhost,1433;Database=Locadora;Uid=sa;Pwd=ELetro66573759000171;Connection Timeout=30;"
End Sub

Private Sub mnu_cad_produtos_novo_Click()
    FrmComCadProdutosNovo.Show
End Sub

Private Sub mnu_cap_consulta_Click()
    FrmComCapConsulta.Show
End Sub

Private Sub mnu_cap_lanc_Click()
    FrmComCapLanc.Show
End Sub

Private Sub mnu_car_consulta_Click()
    FrmComCarConsulta.Show
End Sub

Private Sub mnu_car_lanc_Click()
    FrmComCarLanc.Show
End Sub

Private Sub mnu_emissoes_aluguel_Click()
    FrmComEmissaoAluguel.Show
End Sub

Private Sub mnu_emissoes_consulta_Click()
    FrmComEmissaoConsulta.Show
End Sub

Private Sub mnu_emissoes_cupons_consulta_Click()
    FrmComEmissaoCuponsConsulta.Show
End Sub

Private Sub mnu_emissoes_cupons_emissao_Click()
    FrmComEmissaoCuponsEmissao.Show
End Sub

Private Sub mnu_emissoes_devolucao_Click()
    FrmComEmissaoDevolucao.Show
End Sub

Private Sub mnu_relat_cap_Click()
    FrmComRelatCAP.Show
End Sub

Private Sub mnu_relat_car_Click()
    FrmComRelatCAR.Show
End Sub

Private Sub mnu_relat_cli_Click()
    FrmComRelatClientes.Show
End Sub

Private Sub mnu_relat_ped_Click()
    FrmComRelatPedidos.Show
End Sub

Private Sub mnu_relat_prod_Click()
    FrmComRelatProdutos.Show
End Sub

Private Sub mnu_Sair_Click()
    Unload Me
End Sub
Private Sub mnu_cad_clientes_click()
    FrmComCadClientes.Show
End Sub
Private Sub mnu_cad_produtos_entradas_click()
    FrmComCadProdutosEnt.Show
End Sub
Private Sub mnu_cad_forn_click()
    FrmComCadForn.Show
End Sub
Private Sub mnu_cad_func_click()
    FrmComCadFunc.Show
End Sub


