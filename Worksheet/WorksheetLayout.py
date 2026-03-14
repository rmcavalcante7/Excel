import copy
from typing import Any
import inspect
from openpyxl.utils import column_index_from_string


class WorkSheetLayout:
    chave_principal = 'chave_principal'
    col_origem = 'col_origem'
    col_final = "col_final"
    nome_inicial = 'nome_inicial'
    nome_final = 'nome_final'
    valor = 'valor'
    chave_ligacao = 'chave_ligacao'
    pattern_fill = 'pattern_fill'
    validar = 'validar'

    def __init__(self,
                 estrutura: dict,
                 header_row: int,
                 data_row: int):
        """Classe utilizada para definir os campos da base de dados.
            :param chave_principal: Chave utilizada para identificar cada linha da base
            :param columns_row_number: Número da linha onde estão localizados os nomes das colunas
            :param data_row_number: Número da linha onde inicia os dados da base"""
        self.estrutura = estrutura
        self.header_row = header_row
        self.data_row =  data_row
        # self._data = {}

    def copyEstrutura(self):
        return copy.deepcopy(self.estrutura)  # self.estrutura.copy()

    def getColOrigem(self, nome_campo: str) -> str:
        """Retorna a coluna de origem de um campo.
        :param nome_campo: Nome do campo
        :return: Coluna de origem do campo"""
        if nome_campo not in self.estrutura:
            raise InvalidColumnName(col_name=nome_campo, classe=self.__class__.__name__,
                                    metodo=inspect.currentframe().f_code.co_name)
        return self.estrutura[nome_campo][self.col_origem]

    def getColFinal(self, nome_campo: str) -> str:
        """Retorna a coluna final de um campo.
        :param nome_campo: Nome do campo
        :return: Coluna final do campo"""
        if nome_campo not in self.estrutura:
            raise InvalidColumnName(col_name=nome_campo, classe=self.__class__.__name__,
                                    metodo=inspect.currentframe().f_code.co_name)
        return self.estrutura[nome_campo][self.col_final]

    def getOriginColumnIndex(self, nome_campo: str) -> int|None:
        """Retorna o índice da coluna de origem de um campo.
        :param nome_campo: Nome do campo
        :return: Índice da coluna de origem do campo"""
        col_origem = self.getColOrigem(nome_campo)
        if col_origem:
            return column_index_from_string(col_origem) - 1
        return None

    def getFinalColumnIndex(self, nome_campo: str) -> int|None:
        """Retorna o índice da coluna de origem de um campo.
        :param nome_campo: Nome do campo
        :return: Índice da coluna de origem do campo"""
        col_origem = self.getColFinal(nome_campo)
        if col_origem:
            return column_index_from_string(col_origem) - 1
        return None

class InvalidColumnName(Exception):
    def __init__(self, col_name: str, classe: str = '', metodo: str = ''):
        self.message = f"Coluna informada não existe: '{col_name}'. "\
                       f"Classe '{classe}'.\n" \
                       f"Método '{metodo}()'"
        super().__init__(self.message)

    # @staticmethod
    # def standartStructure(col_origem: str,
    #                       col_final: str,
    #                       nome_inicial: str,
    #                       nome_final: str,
    #                       validar: bool,
    #                       valor: Any=None,
    #                       chave_ligacao: dict=None,
    #                       pattern_fill: dict=None,
    #                       ) -> dict:
    #     """Retorna uma estrutura padronizada de campo.
    #     :param col_origem: Coluna de origem da informação
    #     :param col_final: Coluna de destino da informação,
    #                       caso a informação precise ser transferida para outra planilha
    #     :param nome_inicial: Nome inicial do campo, utilizado para identificar o campo na planilha
    #     :param nome_final: Nome do campo, caso ele tenha que ser transferido para outra planilha
    #     :param validar: Indica se o campo deve ser validado
    #     :param valor: Valor padrão do campo, caso ele tenha um valor padrão
    #     :param chave_ligacao: Chave de ligação entre o campo e o campo de outro layout
    #     :param pattern_fill: Padrão de preenchimento do campo
    #     :return: Estrutura padronizada do campo
    #     """
    #     return {
    #         WorkSheetLayout.col_origem: col_origem,
    #         WorkSheetLayout.col_final: col_final,
    #         WorkSheetLayout.nome_inicial: nome_inicial,
    #         WorkSheetLayout.nome_final: nome_final,
    #         WorkSheetLayout.validar: validar,
    #         WorkSheetLayout.valor: valor,
    #         WorkSheetLayout.chave_ligacao: chave_ligacao,
    #         WorkSheetLayout.pattern_fill: pattern_fill,
    #     }
    #
    # def addField(self,
    #              nome_campo: str,
    #              col_origem: str,
    #              col_final: str,
    #              nome_inicial: str,
    #              nome_final: str,
    #              validar: bool = True,
    #              valor: Any=None,
    #              chave_ligacao: dict=None,
    #              pattern_fill: dict=None
    #              ) -> None:
    #     """Adiciona um novo campo à estrutura de dados.
    #     :param nome_campo: Nome do campo
    #     :param col_origem: Coluna de origem
    #     :param col_final: Coluna de destino
    #     :param nome_inicial: Nome do campo
    #     :param nome_final: Nome do campo
    #     :param valor: Valor do campo
    #     :param status: Status do campo
    #     :param chave_ligacao: Chave de ligação entre os campos de diferentes layouts
    #     :param pattern_fill: Padrão de preenchimento do campo
    #     :param validar: Indica se o campo deve ser validado
    #     """
    #     self._data[nome_campo] = self.standartStructure(
    #         col_origem=col_origem,
    #         col_final=col_final,
    #         nome_inicial=nome_inicial,
    #         nome_final=nome_final,
    #         validar=validar,
    #         valor=valor,
    #         chave_ligacao=chave_ligacao,
    #         pattern_fill=pattern_fill,
    #     )
    #
    # def get_data(self):
    #     # return self._data.copy()
    #     return copy.deepcopy(self._data)  # self.estrutura.copy()
    #
    # @property
    # def data(self):
    #     return self._data