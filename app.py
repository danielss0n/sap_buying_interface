import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from datetime import datetime
import win32com.client
import json
import os
import re
import time

DIR_DB_CENTRO_CUSTOS = "Q:\\EXEMPLO\\app\\databases\\bd_centro_de_custo.txt"
DIR_DB_MTERIAIS = "Q:\\Exemplo\\app\\databases\\materiais.json"
DIR_DB_FORNECEDORES = "Q:\\EXEMPLO\\app\\databases\\fornecedores.json"
DIR_DB_MEUS_PARAMETROS = "Q:\\EXEMPLO\\app\\databases\\parametros_usuario.json"

"""
Classe RequisicaoCotacaoApp(): menu principal
 
Classe ConfiguracoesApp(): Menu de configurações

"""

class DataBase():
    @staticmethod
    def save_user_data(data):
        try:
            with open(DIR_DB_MEUS_PARAMETROS, 'r') as file:
                all_data = json.load(file)
        except FileNotFoundError:
            all_data = []
        except json.JSONDecodeError:
            all_data = []
        
        updated = False
        for item in all_data:
            if item.get('login') == data.get('login'):
                item.update(data)
                updated = True
                break
        
        if not updated:
            all_data.append(data)
        
        with open(DIR_DB_MEUS_PARAMETROS, 'w') as file:
            json.dump(all_data, file, indent=4)


    def carregar_centro_custo():
        linhas = []
        with open(DIR_DB_CENTRO_CUSTOS, 'r', encoding='utf-8') as arquivo:
            linhas = arquivo.readlines()
        return [linha.strip() for linha in linhas]
    

    def carregar_materiais():
        try:
            with open(DIR_DB_MTERIAIS, 'r', encoding='utf-8') as file:
                dados = json.load(file)
        except FileNotFoundError:
            print(f"O arquivo {DIR_DB_MTERIAIS} não foi encontrado.")
            return []
        except json.JSONDecodeError:
            print(f"O arquivo {DIR_DB_MTERIAIS} contém erros no formato JSON.")
            return []

        resultado = []
        for item in dados:
            material = item.get('material', '')
            descricao = item.get('descricao', '')
            resultado.append(f"{material} - {descricao}")

        return resultado
    
    def pegar_dados_do_login(login):
        data = []
        with open(DIR_DB_MEUS_PARAMETROS, 'r') as file:
            data = json.load(file)
        for item in data:
            if item.get('login') == login:
                return item
        return None
    
    def carregar_fornecedores():
        try:
            with open(DIR_DB_FORNECEDORES, 'r') as file:
                return json.load(file)
        except FileNotFoundError:
            return []

    def save_data(bd_dir, data):
        with open(bd_dir, 'w') as file:
            json.dump(data, file, indent=4)

class StringMethods():
    def __init__(self) -> None:
        pass

    def set_mobilizacao(self, mobilizacao:int) -> str:
        if (mobilizacao == 1): 
            return "A"
        if (mobilizacao == 0): 
            return"K"
        
    def get_first_part(self, text:str) -> str:
        return str.split(text, " ")[0]
    
class ME51N():
    def __init__(self):
        self.session = None
        self.connect_sap()
        self.grid = "/app/con[0]/ses[0]/wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
        self.cc_page = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6"
        self.cc_grid = "/app/con[0]/ses[0]/wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/"
        self.textos_page = "wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13"
        self.textos_area = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/"

    def connect_sap(self):
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)
        self.session = connection.Children(0)

    def enter_page(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nme51n"
        self.session.findById("wnd[0]").sendVKey(0)

    def gravar(self):
        self.session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[11]").press()

    def pegar_mensagem_sap(self) -> str:
        return self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").Text
        
    def select_cotation(self):
        self.session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").setFocus()
        self.session.findById("/app/con[0]/ses[0]/wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZNB"

    def grid_set_mobilizacao(self, letter):
        self.session.findById(self.grid).modifyCell(0, "KNTTP", letter)

    def grid_set_material(self, material):
        self.session.findById(self.grid).modifyCell(0, "MATNR", material)

    def grid_set_texto_material(self, texto_material):
        self.session.findById(self.grid).modifyCell(0, "TXZ01", texto_material)

    def grid_set_quantidade(self, quantidade):
        self.session.findById(self.grid).modifyCell(0, "MENGE", quantidade)

    def grid_set_unidade(self, unidade):
        self.session.findById(self.grid).modifyCell(0, "MEINS", unidade)

    def grid_set_gcm(self, gcm):
        self.session.findById(self.grid).modifyCell(0, "EKGRP", gcm)

    def grid_set_data_remessa(self, data):
        self.session.findById(self.grid).modifyCell(0, "EEIND", data)

    def grid_set_centro(self, centro):
        self.session.findById(self.grid).modifyCell(0, "NAME1", centro)

    def grid_set_requisitante(self, requisitante):
        self.session.findById(self.grid).modifyCell(0, "AFNAM", requisitante)

    def grid_press_enter(self):
        self.session.findById(self.grid).pressEnter()
    
    def cc_enter_page(self):
        self.session.findById(self.cc_page).select()
    
    def cc_grid_set_pto_descarga(self, local):
        self.session.findById(f"{self.cc_grid}txtMEACCT1100-ABLAD").text = local
    
    def cc_grid_set_recebedor(self, recebedor):
        self.session.findById(f"{self.cc_grid}txtMEACCT1100-WEMPF").text = recebedor

    def cc_grid_set_conta_razao(self, conta):
        self.session.findById(f"{self.cc_grid}ctxtMEACCT1100-SAKTO").text = conta

    def cc_grid_set_centro_custo(self, centro):
        self.session.findById(f"{self.cc_grid}subKONTBLOCK:SAPLKACB:9002/ctxtCOBL-KOSTL").text = centro
    
    def textos_enter_page(self):
        self.session.findById(self.textos_page).select()

    def textos_set_texto_compra(self, texto):
        self.session.findById(f"{self.textos_area}cntlTEXT_TYPES_0200/shell").selectedNode = "B01"
        self.session.findById(f"{self.textos_area}subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").text = texto

    def textos_set_compra_remessa(self, texto):
        self.session.findById(f"{self.textos_area}cntlTEXT_TYPES_0200/shell").selectedNode = "B03"
        self.session.findById(f"{self.textos_area}subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").text = texto

    def select_layout(self):
        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        self.session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").selectContextMenuItem("&LOAD")
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 342
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 252
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 244
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "252"
        self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

class ME41():
    def __init__(self):
        self.session = None
        self.connect_sap()
        self.grid = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
        self.cc_page = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6"
        self.cc_grid = "/app/con[0]/ses[0]/wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/"
        self.textos_page = "wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13"
        self.textos_area = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/"

    def connect_sap(self):
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)
        self.session = connection.Children(0)
    
    def enter_page(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nME41"
        self.session.findById("wnd[0]").sendVKey(0)

    def gravar(self):
        self.session.findById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[11]").press()

    def pegar_mensagem_sap(self) -> str:
        return self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").Text
    
    def set_tipo_solicitacao(self, tipo):
        self.session.findById("wnd[0]/usr/ctxtRM06E-ASART").text = tipo
            
    def set_prazo_apresentacao(self, data):
        self.session.findById("wnd[0]/usr/ctxtEKKO-ANGDT").text = data
            
    def set_organizacao_compras(self, organizacao):
        self.session.findById("wnd[0]/usr/ctxtEKKO-EKORG").text = organizacao
            
    def set_grupo_compradores(self, grupo):
        self.session.findById("wnd[0]/usr/ctxtEKKO-EKGRP").text = grupo
                    
    def set_organizacao_compras(self, organizacao):
        self.session.findById("wnd[0]/usr/ctxtEKKO-EKORG").text = organizacao

    def clicar_ref_a_req(self):
        self.session.findById("wnd[0]/tbar[1]/btn[27]").press()

    def enviar_requisicao(self, requisicao):
        self.session.findById("wnd[1]/usr/ctxtEKET-BANFN").text = requisicao
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

    def selecionar_linha(self, linha):
        tabela = self.session.findById("wnd[0]/usr/tblSAPMM06ETC_0125")
        tabela.getAbsoluteRow(linha).selected = True

    def clicar_aceitar_detalhe(self):
        self.session.findById("wnd[0]/tbar[1]/btn[16]").press()

    def tecla_enter(self):
        self.session.findById("wnd[0]").sendVKey(0)

    def selecionar_linha_detalhe(self, linha):
        tabela = self.session.findById("wnd[0]/usr/tblSAPMM06ETC_0320")
        tabela.getAbsoluteRow(linha).selected = True
        self.session.findById("wnd[0]/tbar[1]/btn[7]").press()

    def escrever_fornecedor(self, fornecedor):
        self.session.findById("wnd[0]/usr/ctxtEKKO-LIFNR").text = fornecedor

class RequisicaoCotacaoService():
    def __init__(self, requisicao):
        self.requisicao = requisicao
        self.setar_variaveis()

    def setar_variaveis(self) -> None:
        self.requisitante = self.requisicao["requisitante"]
        self.mobilizacao = self.requisicao["mobilizacao"]
        self.unidade = self.requisicao["unidade"]
        self.material = self.requisicao["material"]
        self.material_descricao = self.requisicao["material_descricao"]
        self.compra_descricao = self.requisicao["compra_descricao"]
        self.quantidade = self.requisicao["quantidade"]
        self.centro = self.requisicao["centro"]
        self.pto_descarga = self.requisicao["pto_descarga"]
        self.centro_de_custo = self.requisicao["centro_de_custo"]
        self.gcm = self.requisicao["gcm"]
        self.conta_razao = self.requisicao["conta_razao"]

    def validar_dados(self) -> tuple[bool, str]:
        if not isinstance(self.requisicao, dict):
            return False, "Erro de sistema: A requisição deve ser um dicionário"

        if not self.validar_vazio(self.requisitante):
            return False, "Campo de Requisitante inválido:\nValor não pode ser vazio"

        if not self.validar_tamanho_igual(self.mobilizacao, 1):
            return False, "Campo de Mobilização inválida:\nValor tem mais de um caractere"

        if not self.validar_unidade(self.unidade):
            return False, "Campo de Unidade inválida:\nValor não permitido"
        
        if not self.validar_numerico(self.material):
            return False, "Campo de Material inválido:\nValor não numérico"

        if not self.validar_tamanho_menor(self.material_descricao, 40):
            return False, "Campo de Descrição do material inválido:\nExcede 40 caracteres"
        
        if not self.validar_vazio(self.material_descricao):
            return False, "Campo de Descrição do material inválido:\nEstá vazio"

        if not self.validar_tamanho_menor(self.compra_descricao, 500):
            return False, "Campo de Descrição da compra inválida:\nExcede 500 caracteres"
        
        if not self.validar_vazio(self.compra_descricao):
            return False, "Campo de Descrição da compra inválida:\nEstá vazio"

        if not self.validar_numerico(self.quantidade):
            return False, "Campo de Quantidade inválida:\nValor não numérico"

        if not self.validar_numerico(self.centro):
            return False, "Campo de Centro inválido:\nEstá vazio"

        if not self.validar_vazio(self.pto_descarga):
            return False, "Campo de Ponto de descarga inválido:\nEstá vazio"

        if not self.validar_numerico(self.centro_de_custo):
            return False, "Campo de Centro de custo inválido:\nValor não numérico"

        if not self.validar_vazio(self.gcm):
            return False, "Campo de GCM inválido:\nEstá vazio"

        if not self.validar_numerico(self.conta_razao):
            return False, "Campo de Conta razão inválida:\nValor não não é numérico"

        return True, "Sucesso ao validar dados da requisição"

    def validar_numerico(self, valor:str) -> bool:
        return valor.isnumeric()

    def validar_vazio(self, requisitante:str) -> bool:
        return bool(requisitante and requisitante.strip())

    def validar_tamanho_igual(self, valor:str, tamanho:int) -> bool:
        return valor is not None and len(valor) == tamanho

    def validar_tamanho_menor(self, material_descricao:str, tamanho:int) -> bool:
        return len(material_descricao) <= tamanho
    
    def validar_unidade(self, unidade:str) -> bool:
        return unidade in {"SAC", "CE", "CJ", "JG", "KG", "L", "M", "M3", "PEÇ", "MIL", "UN"}      

class SapService():
    def __init__(self):
        pass

    def controller_requisicao_cotacao(self, data) -> str:
        from app.sap.SAP_ME51N import ME51N

        sap = ME51N()
        sap.enter_page()
        sap.select_cotation()
        sap.select_layout()
        sap.grid_set_material(data["material"])
        sap.grid_set_mobilizacao(data["mobilizacao"])
        sap.grid_set_texto_material(data["material_descricao"])
        sap.grid_set_quantidade(data["quantidade"])
        sap.grid_set_unidade(data["unidade"])
        sap.grid_set_gcm(data["gcm"])
        sap.grid_set_data_remessa(data["data_atual"])
        sap.grid_set_centro(data["centro"])
        sap.grid_set_requisitante(data["requisitante"])
        sap.grid_press_enter()
        sap.cc_grid_set_conta_razao(data["conta_razao"])
        sap.cc_grid_set_pto_descarga(data["pto_descarga"])
        sap.cc_grid_set_recebedor(data["requisitante"])
        sap.cc_grid_set_centro_custo(data["centro_de_custo"])
        sap.textos_enter_page()
        sap.textos_set_texto_compra(data["compra_descricao"])
        sap.textos_set_compra_remessa(data["compra_descricao"])

        time.sleep(3) # Tempo de processamento do SAP até mostrar a mensagem
        
        # req_cotacao = sap.pegar_mensagem_sap()
        req_cotacao = "Requisição de cotação criada sob nº 3000237644"

        return self.extrair_numero(req_cotacao)

    def controller_envio_fornecedores(self, data):
        from app.sap.SAP_ME41 import ME41

        data = json.loads(data)

        me41 = ME41()
        me41.enter_page()
        me41.set_tipo_solicitacao("ZME")
        me41.set_prazo_apresentacao(data["data_prazo"])
        me41.set_grupo_compradores("A22")
        me41.set_organizacao_compras("1001")
        me41.clicar_ref_a_req()
        me41.enviar_requisicao(data["requisicao"])
        me41.selecionar_linha(0)
        me41.clicar_aceitar_detalhe()
        me41.tecla_enter()
        me41.tecla_enter()
        me41.tecla_enter()
        me41.selecionar_linha_detalhe(0)

        fornecedores = data["fornecedores"]
        for fornecedor in fornecedores:
            me41.escrever_fornecedor(fornecedor)
            me41.tecla_enter()
        #me41.gravar()

    def extrair_numero(self, string):
        padrao = r'\b\d{10}\b'
        return re.search(padrao, string).group()

class FornecedorApp():
    def __init__(self):
        root = tk.Tk()
        self.root = root
        self.root.title("Cadastrar Fornecedores")
        self.root.geometry("400x400")
        
        self.group_options = ["CALDEIRARIA", "FERRAMENTAS", "INFORMATICA", "MATERIAIS AUDIOVISUAL", "OUTROS", "USINAGEM"]
        self.DIR_DB_FORNECEDORES = DIR_DB_FORNECEDORES 

        self.tree = ttk.Treeview(self.root, columns=("Grupo", "Número", "Nome"), show='headings')
        self.update_window()
        self.update_table()

        # Botões
        self.btn_add = tk.Button(self.root, text="Adicionar Item", command=self.add_item)
        self.btn_add.pack(pady=10)
        self.btn_remove = tk.Button(self.root, text="Remover Item", command=self.remove_item)
        self.btn_remove.pack(pady=10)
        root.mainloop()

    
    def load_data(self):
        try:
            with open(DIR_DB_FORNECEDORES, 'r') as file:
                return json.load(file)
        except FileNotFoundError:
            return []


    def save_data(self, data):
        with open(DIR_DB_FORNECEDORES, 'w') as file:
            json.dump(data, file, indent=4)


    def update_table(self):
        data = self.load_data()
        data = sorted(data, key=lambda x: x['group'])
        
        # Limpar a Treeview
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Adicionar novos dados
        for item in data:
            self.tree.insert("", tk.END, values=(item['group'], item['number'], item['nome']))


    def add_item(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Adicionar Item")
        add_window.geometry("300x300")

        tk.Label(add_window, text="Grupo:").pack(pady=5)
        group_combobox = ttk.Combobox(add_window, values=self.group_options, width=25)
        group_combobox.pack(pady=5)
        group_combobox.set("Selecione um grupo")

        tk.Label(add_window, text="Número:").pack(pady=5)
        number_entry = tk.Entry(add_window, width=25)
        number_entry.pack(pady=5)

        tk.Label(add_window, text="Nome:").pack(pady=5)
        name_entry = tk.Entry(add_window, width=25)
        name_entry.pack(pady=5)

        def save_new_item():
            group = group_combobox.get()
            number = number_entry.get()
            name = name_entry.get()
            
            if not number.isnumeric():
                messagebox.showinfo("Erro", "Número deve ser numérico!")
                return

            if group and number and name:
                data = self.load_data()
                data.append({"group": group, "number": number, "nome": name})
                self.save_data(data)
                self.update_table()
                add_window.destroy()

        tk.Button(add_window, text="Salvar", command=save_new_item).pack(pady=10)


    def remove_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item = self.tree.item(selected_item)
            item_values = item['values']
            number = str(item_values[1])
            data = self.load_data()
            dados = [item for item in data if item.get('number') != number]

            self.save_data(dados)
            self.update_table()


    def update_window(self):
        self.tree.heading("Grupo", text="Grupo")
        self.tree.heading("Número", text="Número")
        self.tree.heading("Nome", text="Nome")
        self.tree.column("Grupo", width=100)
        self.tree.column("Número", width=100)
        self.tree.column("Nome", width=100)
        self.tree.pack(expand=True, fill='both')

class MateriaisApp():
    def __init__(self):
        root = tk.Tk()
        self.root = root
        self.root.title("Cadastrar Materiais")
        self.root.geometry("400x400")
        
        self.DIR_DB_FORNECEDORES = DIR_DB_MTERIAIS 

        self.tree = ttk.Treeview(self.root, columns=("material", "descricao"), show='headings')
        self.update_window()
        self.update_table()

        # Botões
        self.btn_add = tk.Button(self.root, text="Adicionar Item", command=self.add_item)
        self.btn_add.pack(pady=10)
        self.btn_remove = tk.Button(self.root, text="Remover Item", command=self.remove_item)
        self.btn_remove.pack(pady=10)
        root.mainloop()

    
    def load_data(self):
        try:
            with open(DIR_DB_MTERIAIS, 'r') as file:
                return json.load(file)
        except FileNotFoundError:
            return []


    def save_data(self, data):
        with open(DIR_DB_MTERIAIS, 'w') as file:
            json.dump(data, file, indent=4)


    def update_table(self):
        data = self.load_data()
        data = sorted(data, key=lambda x: x['descricao'])
        
        # Limpar a Treeview
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Adicionar novos dados
        for item in data:
            self.tree.insert("", tk.END, values=(item['material'], item['descricao']))


    def add_item(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Adicionar Item")
        add_window.geometry("300x300")

        tk.Label(add_window, text="Número:").pack(pady=5)
        entry_material = tk.Entry(add_window, width=25)
        entry_material.pack(pady=5)

        tk.Label(add_window, text="Descrição:").pack(pady=5)
        entry_descricao = tk.Entry(add_window, width=25)
        entry_descricao.pack(pady=5)

        def save_new_item():
            material = entry_material.get()
            descricao = entry_descricao.get()
            
            if not material.isnumeric():
                messagebox.showinfo("Erro", "Material deve ser numérico!")
                return
            
            if not descricao:
                messagebox.showinfo("Erro", "Deve conter descrição!")
                return

            if material and descricao:
                data = self.load_data()
                data.append({"material": material, "descricao": descricao})
                self.save_data(data)
                self.update_table()
                add_window.destroy()

        tk.Button(add_window, text="Salvar", command=save_new_item).pack(pady=10)


    def remove_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item = self.tree.item(selected_item)
            item_values = item['values']
            number = str(item_values[0])
            data = self.load_data()
            dados = [item for item in data if item.get('material') != number]

            self.save_data(dados)
            self.update_table()


    def update_window(self):
        self.tree.heading("material", text="Material")
        self.tree.heading("descricao", text="Descrição")
        self.tree.column("material", width=100)
        self.tree.column("descricao", width=100)
        self.tree.pack(expand=True, fill='both')

class MeuParametrosApp():
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Meus parâmetros")
        self.root.geometry("400x400")

        self.user_login = os.getenv('USERNAME')
        self.user_data = DataBase.pegar_dados_do_login(self.user_login)

        tk.Label(self.root, text="Conta razão:").pack(pady=5)
        self.entry_razao = tk.Entry(self.root, width=25)
        self.entry_razao.pack(pady=5)
        self.entry_razao.insert(0, self.user_data["conta_razao_padrao"])

        tk.Label(self.root, text="Local entrega:").pack(pady=5)
        self.entry_local_entrega = tk.Entry(self.root, width=25)
        self.entry_local_entrega.pack(pady=5)
        self.entry_local_entrega.insert(0, self.user_data["local_entrega_padrao"])

        tk.Label(self.root, text="Grupo comprador:").pack(pady=5)
        self.entry_grupo_comprador = tk.Entry(self.root, width=25)
        self.entry_grupo_comprador.pack(pady=5)
        self.entry_grupo_comprador.insert(0, self.user_data["grupo_comprador_padrao"])

        # Botões
        self.btn_add = tk.Button(self.root, text="Salvar", command=self.update)
        self.btn_add.pack(pady=10)

        self.root.mainloop()

    def update(self):
        data = {
            "login": self.user_login,
            "conta_razao_padrao": self.entry_razao.get(),
            "local_entrega_padrao": self.entry_local_entrega.get(),
            "grupo_comprador_padrao": self.entry_grupo_comprador.get()
        }
        DataBase.save_user_data(data)
        self.root.destroy()

class EnvioFornecedoresApp():
    def __init__(self, requisicao):
        self.root = tk.Tk()
        self.root.title("Lista de Checkboxes")
        self.requisicao = requisicao

        tk.Label(self.root, text="Prazo desejado (dd.mm.aaaa):", font=('Arial', 10), fg='grey').grid(row=0, column=0, padx=10, sticky='w')

        self.entry_prazo = tk.Entry(self.root, width=20, font=('Arial', 10))
        self.entry_prazo.grid(row=0, column=1, padx=10, pady=(10, 5), sticky='w')

        tk.Label(self.root, text="Escolha o tipo", font=('Arial', 10), fg='grey').grid(row=2, column=0, padx=10, sticky='w')

        group_options = ["CALDEIRARIA", "FERRAMENTAS", "INFORMATICA", "MATERIAIS AUDIOVISUAL", "OUTROS", "USINAGEM"]
        self.tipos = ttk.Combobox(self.root, state="readonly", font=('Arial', 10), values = group_options, width=20)
        self.tipos.grid(row=2, column=1, pady=(10, 5), sticky='w')
        self.tipos.bind("<<ComboboxSelected>>", self.on_combobox_change)

        button_requisitar = tk.Button(self.root, text="Enviar", command=self.enviar, font=('Arial', 16), width=10)
        button_requisitar.grid(row=0, column=2, columnspan=2, pady=20)

        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        self.root.mainloop()


    def on_combobox_change(self, event):
        selected_value = self.tipos.get()
        print(f"MUDOU PARA: {selected_value}")

        for widget in self.root.grid_slaves(row=4):
            widget.destroy()
        
        forcedores_do_mesmo_tipo = []
        data = DataBase.carregar_fornecedores()
        for item in data:
            if item["group"] == selected_value:
                forcedores_do_mesmo_tipo.append(f"{item['number']} - {item['nome']}")

        self.checkbox_vars, self.checkbox_nomes = self.criar_lista_checkboxes(forcedores_do_mesmo_tipo)

    def criar_lista_checkboxes(self, itens):
        canvas = tk.Canvas(self.root, width=100)
        canvas.grid(row=4, column=1, sticky='nsew')

        scrollbar = tk.Scrollbar(self.root, orient='vertical', command=canvas.yview)
        scrollbar.grid(row=4, column=2, sticky='ns')

        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor='nw')
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        vars = []
        nomes = []

        for i, item in enumerate(itens):
            var = tk.IntVar()
            checkbutton = tk.Checkbutton(frame, text=item, variable=var)
            checkbutton.grid(row=i, column=0, sticky='w', padx=5, pady=2)
            vars.append(var) 
            nomes.append(item) 

        print(vars, nomes)
        return vars, nomes

    def obter_nomes_selecionados(self):
        nomes_selecionados = []
        for checkbox in self.checkbox_vars:
            print(checkbox.get())

        for i, var in enumerate(self.checkbox_vars):
            if var.get() == 1:
                nomes_selecionados.append(self.checkbox_nomes[i])
        return nomes_selecionados


    def enviar(self):
        selecionados = self.obter_nomes_selecionados()

        print(selecionados)

        data = self.entry_prazo.get()

        print(data)

        if not data or data == "":
            messagebox.showerror("Erro", "A data deve ser informada")
            return

        if len(selecionados) < 1:
            messagebox.showerror("Erro", "Escolha pelo menos 1 fornecedor")
            return

        fornecedores_numeros = [fornecedor.split(" ")[0] for fornecedor in selecionados]
        for fornecedor in selecionados:
            fornecedores_numero = fornecedor.split(" ")[0]
            fornecedores_numeros.append(fornecedores_numero)

        item_json = {
            "requisicao": self.requisicao,
            "data_prazo": data,
            "fornecedores": fornecedores_numeros
        }
        
        dados = json.dumps(item_json, ensure_ascii=False, indent=4)
        sap = SapService()
        sap.controller_envio_fornecedores(dados)

class ConfiguracoesApp():
    def __init__(self):
      # Inicializa a janela principal usando customtkinter
        self.root = ctk.CTk()
        self.root.title("Configurações")
        self.root.geometry("400x300")

        # Configure o tema do customtkinter
        ctk.set_appearance_mode("Dark")  # ou "Light"
        ctk.set_default_color_theme("blue")  # escolha uma cor de tema ou defina seu próprio tema

        # Criação dos botões com customtkinter
        button_fornecedores = ctk.CTkButton(self.root, text="Fornecedores", command=self.fornecedores, font=('Arial', 16), width=200)
        button_fornecedores.grid(row=0, column=0, columnspan=2, pady=20)

        button_materiais = ctk.CTkButton(self.root, text="Materiais", command=self.materiais, font=('Arial', 16), width=200)
        button_materiais.grid(row=1, column=0, columnspan=2, pady=20)

        button_param_padrao = ctk.CTkButton(self.root, text="Parâmetros padrão", command=self.parametros_padrao, font=('Arial', 16), width=200)
        button_param_padrao.grid(row=2, column=0, columnspan=2, pady=20)

        # Ajuste do layout para o grid
        self.root.grid_rowconfigure([0, 1, 2], weight=1)
        self.root.grid_columnconfigure([0, 1], weight=1)

        # Inicia o loop principal da aplicação
        self.root.mainloop()
    
    def fornecedores(self):
        FornecedorApp()
        
    def materiais(self):
        MateriaisApp()

    def parametros_padrao(self):
        MeuParametrosApp()

class RequisicaoCotacaoApp:
    def __init__(self):
        # Inicializa a janela principal usando customtkinter
        self.root = ctk.CTk()
        self.root.title("REQUISIÇÃO DE COTAÇÃO")
        self.root.geometry("500x530")  # Ajustado para um tamanho mais apropriado

        # Definindo variáveis
        self.user_login = os.getenv('USERNAME')
        self.user_data = DataBase.pegar_dados_do_login(self.user_login)
        if self.user_data is None:
            # Implementar ação caso não haja dados do usuário
            pass

        self.materiais = DataBase.carregar_materiais()
        self.centros_custo = DataBase.carregar_centro_custo()

        # Configura o frame principal
        self.frame = ctk.CTkFrame(self.root)  # Remove padding, que não é suportado
        self.frame.pack(fill=tk.BOTH, expand=True)

        self.criar_elementos()
        self.root.mainloop()

    def criar_elementos(self):
        # Checkbuttons "Sim" e "Não"
        ctk.CTkLabel(self.frame, text="Imobilizado", font=('Arial', 14)).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        self.var_sim = tk.IntVar()
        self.check_sim = ctk.CTkCheckBox(self.frame, text="Sim", variable=self.var_sim, onvalue=1, offvalue=0, font=('Arial', 14))
        self.check_sim.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)

        # Material
        ctk.CTkLabel(self.frame, text="Escolha o material", font=('Arial', 14)).grid(row=1, column=0, sticky=tk.W, padx=10,  pady=5)
        self.combo_material = ctk.CTkComboBox(self.frame, state="readonly", font=('Arial', 14), values=self.materiais, width=200)
        self.combo_material.grid(row=1, column=1, padx=10,  pady=5, sticky=tk.W)

        # Texto Material
        ctk.CTkLabel(self.frame, text="Digite o texto do material", font=('Arial', 14)).grid(row=2, column=0, padx=10, sticky=tk.W, pady=5)
        self.text_material = ctk.CTkTextbox(self.frame, width=300, height=60, font=('Arial', 14))
        self.text_material.grid(row=2, column=1, padx=10,  pady=5, sticky=tk.W)

        # Texto da compra
        ctk.CTkLabel(self.frame, text="Digite o texto da compra", font=('Arial', 14)).grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        self.text_compra = ctk.CTkTextbox(self.frame, width=300, height=120, font=('Arial', 14))
        self.text_compra.grid(row=3, column=1, padx=10,  pady=5, sticky=tk.W)

        # Quantidade
        ctk.CTkLabel(self.frame, text="Digite a quantidade", font=('Arial', 14)).grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
        self.spinbox = ctk.CTkEntry(self.frame, width=100, font=('Arial', 14))
        self.spinbox.grid(row=4, column=1,  padx=10, pady=5, sticky=tk.W)
        self.spinbox.insert(0, '1')  # Valor padrão

        # Unidade
        ctk.CTkLabel(self.frame, text="Escolha a unidade", font=('Arial', 14)).grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
        self.combo_unidade = ctk.CTkComboBox(self.frame, width=100, font=('Arial', 14), state="readonly", values=["UN", "PEÇ"])
        self.combo_unidade.grid(row=5, column=1, padx=10,  pady=5, sticky=tk.W)

        # Centro de custo
        ctk.CTkLabel(self.frame, text="Escolha o centro de custo", font=('Arial', 14)).grid(row=6, column=0, sticky=tk.W, padx=10, pady=5)
        self.centro_custo = ctk.CTkComboBox(self.frame, state="readonly", font=('Arial', 14), values=self.centros_custo, width=200)
        self.centro_custo.grid(row=6, column=1, padx=10, pady=5, sticky=tk.W)

        # Botão Requisitar Cotação
        button_requisitar = ctk.CTkButton(self.frame, text="Requisitar Cotação", command=self.requisitar_cotacao, font=('Arial', 16), width=200)
        button_requisitar.grid(row=7, column=0, padx=10, columnspan=2, pady=20)

        # Botão Requisitar Configurações
        button_config = ctk.CTkButton(self.frame, text="Configurações", command=self.mostrar_config, font=('Arial', 16), width=200)
        button_config.grid(row=8, column=0, padx=10,  columnspan=2, pady=20)

    def mostrar_config(self):
        ConfiguracoesApp()
        
    def requisitar_cotacao(self):
        if self.var_sim.get() == 1:
            messagebox.showerror("Alerta!!!!!", "Sistema para itens imobilizados ainda não está pronto")
            return
        
        strings = StringMethods()
        requisicao = {
            "requisitante": os.getenv('USERNAME'),
            "mobilizacao": strings.set_mobilizacao(self.var_sim.get()),
            "material": strings.get_first_part(self.combo_material.get()),
            "material_descricao": self.text_material.get("1.0", tk.END).strip(),
            "compra_descricao": self.text_compra.get("1.0", tk.END).strip(),
            "quantidade": self.spinbox.get(),
            "unidade": self.combo_unidade.get(),
            "data_atual": datetime.now().strftime('%d.%m.%Y'),
            "centro": "1200",
            "pto_descarga": "PREDIO 21 CALDEIRARIA",
            "centro_de_custo": strings.get_first_part(self.centro_custo.get()),
            "gcm": "A22",
            "conta_razao": "411010003",
        }

        requisicao_service = RequisicaoCotacaoService(requisicao)
        mensagem_validacao = requisicao_service.validar_dados()
        validacao_passou = mensagem_validacao[0]
        validacao_mensagem = mensagem_validacao[1]

        sap_service = SapService()
        if validacao_passou:
            messagebox.showinfo("Sucesso", validacao_mensagem)
            numero_cotacao = sap_service.controller_requisicao_cotacao(requisicao)
            self.root.destroy()
            EnvioFornecedoresApp(numero_cotacao)
        
        if not validacao_passou:
            messagebox.showinfo("Falha", validacao_mensagem)

    def create_user(self):
        data = {
            "login": self.user_login,
            "conta_razao_padrao": "411010003",
            "local_entrega_padrao": "PREDIO 22 CALDEIRARIA",
            "grupo_comprador_padrao": "A22"
        }
        DataBase.save_user_data(data)
        self.root.destroy()

    def combobox_load_values(self):
        self.combo_material['values'] = DataBase.carregar_materiais()
        self.centro_custo['values'] = DataBase.carregar_centro_custo()

RequisicaoCotacaoApp()
