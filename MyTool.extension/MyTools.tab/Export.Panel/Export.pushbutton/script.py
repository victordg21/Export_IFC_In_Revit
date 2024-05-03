# -- coding: utf-8 --

import os
import clr
from pyrevit import DB, forms, revit 
from Autodesk.Revit.DB import Transaction, IFCExportOptions
import ctypes
from ctypes import wintypes

def find_download_folder(): # Função para levar o IFC a pasta de Downloads quando exportado

    CSIDL_PERSONAL = 5 # Constante para encontrar a pasta "Meus Documentos" que corresponde ao número 5
    SHGFP_TYPE_CURRENT = 0   

    buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
    # O cytpes especifica o caminho da pasta para "Meus Documentos" com o CSIDL_PERSONAL
    return os.path.join(os.path.dirname(buf.value), 'Downloads')
    # O os.path.join extrai o diretório de "Meus Documentos" a partir do buf.value e junta ele 
    # Com o subdiretorio "Downloads", formando o caminho para guardar o arquivo quando exportado

def define_config_ifc(): 

    export_config_ifc = { # Definindo todos os tipos distintos de configurações para exportar o IFC
        # Obs: Não encontrei todos os tipos no site da documentacao "https://www.revitapidocs.com/2019/03f1bce3-dd39-9deb-c732-db82474cb40b.htm"
        "Vista de coordenação IFC 2x3 2.0": DB.IFCVersion.IFC2x3CV2,                            
        "Vista de coordenação IFC 2x3": DB.IFCVersion.IFC2x3,
        "Vista de entrega IFC 2x3 Basic FM": DB.IFCVersion.IFC2x3BFM,
        "Vista de coordenação IFC 2x2": DB.IFCVersion.IFC2x2,
        "Vista do material de entrega do projeto IFC2x3 COBie 2.": DB.IFCVersion.IFCCOBIE,
        "Vista de referência IFC4": DB.IFCVersion.IFC4RV,
        "Vista de transferência de design IFC4": DB.IFCVersion.IFC4DTV,
    }

    config_option = forms.SelectFromList.show( # Utilizando o pyrevit.forms para dispor as opções em select
        export_config_ifc.keys(),
        multiselect=False,
        title="Configure a versão do IFC a ser exportada",
        button_name="Selecione a opção para exportar"
    )

    if not config_option: # Retornando um alert se o processo for cancelado
        forms.alert("Exportação IFC cancelada.")
        return None

    return export_config_ifc[config_option]

def export_ifc(config_ifc):
    file_path = find_download_folder()

    active_view = revit.doc.ActiveView
    options = IFCExportOptions()
    options.FileVersion = config_ifc # Linkando a versão do arquivo com a versão escolhida na função do define

    with Transaction(revit.doc, 'Export IFC') as t:
        try:
            t.Start()
            revit.doc.Export(file_path, active_view.Name, options) #Passando os dados de rota, opção e nome ativo da exportação
            print('Arquivo IFC exportado com sucesso para ' + str(file_path))
        except Exception as e:
            t.RollBack()
            print("Erro na exportação do arquivo IFC: " + str(e))


def main():

    # Função principal que coordena a seleção da versão do IFC e do caminho de exportação,
    # e inicia o processo de exportação IFC.

    config_ifc = define_config_ifc()
    if not config_ifc:
        print("Ocorreu um erro na hora da seleção da opção a ser exportada")
        return 

    export_ifc(config_ifc)


if __name__ == "__main__":
    main()