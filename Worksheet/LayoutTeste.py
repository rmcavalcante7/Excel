from Projeto.Scripts.Excel.Worksheet.WorkSheet import WorkSheetLayout


class LayoutTESTE(WorkSheetLayout):
    def __init__(self):
        super().__init__(
            chave_principal='Teste1',
            columns_row_number=1,
            data_row_number=2
        )

        self.addField(
            nome_campo='Teste1',
            col_origem='A',
            col_final='A',
            nome_inicial='Teste1',
            nome_final='Teste1',
            value='',
            status=''
        )

        self.addField(
            nome_campo='Teste2',
            col_origem='B',
            col_final='B',
            nome_inicial='Teste2',
            nome_final='Teste2',
            value='',
            status=''
        )


