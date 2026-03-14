import traceback
from typing import Any

from openpyxl.utils import column_index_from_string

from Projeto.Scripts.Auxiliar.String import TratarString
import inspect

from Projeto.Scripts.Excel.Excel import Excel
from Projeto.Scripts.Excel.Worksheet.WorksheetLayout import WorkSheetLayout
from Projeto.Scripts.Excel.Workbook.WorkBook import WorkBook


class WorkSheet:
    def __init__(self,
                 workbook: WorkBook,
                 worksheet_name: str,
                 layout: [WorkSheetLayout | None]):
        self.workbook = workbook
        self.worksheet = self.workbook.readWorkSheet(worksheet_name)
        self.dict_base = {}
        self.layout = layout

    def __repr__(self):
        return (f'Classe: {self.__class__.__name__}\n'
                f'Worksheet Name: ({self.worksheet.title})')

    def validateWorkSheet(self):
        """Função utilizada para validar a estrutura da planilha.
        :return: None
        :raises: InvalidBaseLayout: Caso a estrutura da planilha não esteja de acordo com o padrão esperado.
        :raises: Exception: Caso ocorra algum erro inesperado.
        """
        nome_coluna_esperada = ''
        nome_col_base = ''
        pos_col = ''
        status_list = []
        column_row_number = 'Indefinido'
        try:
            for row_number, row in enumerate(
                    self.worksheet.iter_rows(min_row=self.layout.header_row,
                                             max_row=self.layout.header_row,
                                             values_only=True)):
                for key, values in self.layout.estrutura.items():
                    if not values[WorkSheetLayout.validar]:
                        continue

                    letra_coluna = values[WorkSheetLayout.col_origem].upper()
                    nome_coluna_esperada = ''
                    nome_col_base = ''
                    pos_col = f'{letra_coluna}{self.layout.header_row}'
                    nome_coluna_esperada = TratarString.substituirCaractereEspecial(
                        TratarString.tratarEspacos(values[WorkSheetLayout.nome_inicial])).upper()

                    nome_col_base = TratarString.substituirCaractereEspecial(
                        TratarString.tratarEspacos(
                            str(row[column_index_from_string(letra_coluna) - 1]).upper()
                        )
                    ).upper()

                    if nome_coluna_esperada != nome_col_base:
                        status_list.append(
                            f"A coluna '{nome_coluna_esperada}' não está na posição esperada: '{pos_col}'. "
                            f"A coluna '{nome_col_base}' foi encontrada no lugar dela.")
                break

            if status_list:
                raise InvalidWorkSheetLayout(
                    file_path=self.workbook.file_path,
                    worksheet_name=self.worksheet.title,
                    worksheet_layout=self.layout,
                    lista_erros=status_list,
                    classe=self.__class__.__name__,
                    metodo=inspect.currentframe().f_code.co_name
                )


        except InvalidWorkSheetLayout as ibl:
            raise ibl

        except Exception as e:
            colunas_esperadas = '; '.join([self.layout.estrutura[key][self.layout.nome_inicial]
                                           for key in self.layout.estrutura.keys()])
            raise Exception(
                f'Erro desconhecido no método {inspect.currentframe().f_code.co_name}() da classe {self.__class__.__name__}.\n'
                f'Nome da coluna esperada: "{nome_coluna_esperada}".\n'
                f'Nome da coluna na base: "{nome_col_base}".\n'
                f'Posição esperada da coluna na base: "{pos_col}".\n'
                f'Verifique se o arquivo está no padrão esperado.\n'
                f'Verifique se as colunas do arquivo estão na linha {column_row_number}.\n'
                f'Colunas esperadas: {colunas_esperadas}.\n'
                f'Arquivo em tratativa: "{self.workbook.file_path}"\n'
                f"Erros de layout encontrados:\n{'|'.join(status_list) if status_list else ''}\n"
                f'Exceção:\n{traceback.format_exc()}')

    def writeToWorksheetCell(self, row_number: int, col_name: str, value: Any):
        """Escreve um valor em uma célula específica da planilha.
        :param row_number: Número da linha onde a célula está localizada.
        :param col_name: Nome da coluna onde a célula está localizada.
        :param value: Valor a ser escrito na célula.
        """

        # self.worksheet[f'{self.layout.estrutura[col_name][
        #     self.layout.col_final]}{row_number}'] = value
        self.worksheet[f'{self.layout.getColFinal(col_name)}{row_number}'] = value

    # TODO: fazer um método para ler e indexar a base de dados em função de uma dada chave
    # def getDictBase(self, campo_chave=None) -> Generator[WorkSheetLayout, Any, None]:
    #     work_sheet_layout = self.layout.copyEstrutura()
    #     for field in work_sheet_layout:
    #         if field in self.layout.lista_exececao:
    #             continue
    #         coluna = work_sheet_layout[field][self.layout.col_origem]
    #         valor_celula = row[column_index_from_string(coluna) - 1]
    #         work_sheet_layout[field][self.layout.valor] = valor_celula if valor_celula else ""
    #
    #     # chave = work_sheet_layout[campo_chave][self.worksheet_layout.valor] if campo_chave else linha_id
    #     # row_result[chave] = self.worksheet_layout.data
    #     # self.dict_base[chave] = work_sheet_layout
    #     yield work_sheet_layout






class InvalidWorkSheetLayout(Exception):
    def __init__(self,
                 file_path: str,
                 worksheet_name: str,
                 worksheet_layout: WorkSheetLayout,
                 lista_erros: list,
                 classe: str = '',
                 metodo: str = ''
                 ):
        colunas_esperadas = '\n'.join([(f"Coluna '{worksheet_layout.estrutura[col]['nome_inicial']}' na posição "
                              f"'{worksheet_layout.estrutura[col]['col_origem']}{worksheet_layout.header_row}'.")
                             for col in worksheet_layout.estrutura.keys()])

        self.message = f"A aba '{worksheet_name}' da base fixa '{file_path}' não está no padrão esperado.\n" \
                       f"Padrão Esperado:\n{colunas_esperadas}\n\n" \
                       f"Erros encontrados:\n{''.join(lista_erros)}\n" \
                       f"Classe '{classe}'.\n" \
                       f"Método '{metodo}()'"
        super().__init__(self.message)

class EmptyBase(Exception):
    def __init__(self,
                 file_path: str,
                 worksheet_name: str,
                 classe: str = '',
                 metodo: str = ''
                 ):
        self.message = (
            f"Base '{worksheet_name}' está vazia.\n"
            f"Verifique se não houve alterações no padrão da aba '{worksheet_name}' da base fixa "
            f"'{file_path}'."
            f"Classe '{classe}'.\n"
            f"Método '{metodo}()'"
        )
        super().__init__(self.message)


if '__main__' == __name__:
    from Projeto.Scripts.Excel.Worksheet.LayoutTeste import LayoutTESTE
    from pprint import pprint

    layout_teste = LayoutTESTE()
    file_path_teste = f'Teste.xlsx'
    planilha = WorkBook(excel=Excel(),
                        file_path=file_path_teste,
                        read_only=True,
                        data_only=True)
    worksheet_bloqueio_name = 'Teste'

    # Criação do objeto WorkSheet
    work_sheet = WorkSheet(workbook=planilha, worksheet_name=worksheet_bloqueio_name, layout=layout_teste)

    work_sheet.validateWorkSheet()
    work_sheet.getDictBase()
    pprint(work_sheet.dict_base)
    #