import copy 
class BaseSheet:
    def __init__(self, ws, sheet_config, input_data):
        self.ws = ws
        self.cfg = sheet_config
        self.input = input_data

    def get_last_filled_column(self):
        start_clmn = self.cfg["start_column"]
        start_year = self.cfg["start_year"]
        #start_month = self.cfg["start_month"]
        input_year = self.input["report_period"]["year"]
        input_month = int(self.input["report_period"]["date_start"].month)

        if input_month != 1:
            current_clmn = (start_clmn + input_month) - 2
        else:
            current_clmn = start_clmn - 1
        
        while start_year < input_year:
            input_year -= 1
            current_clmn += 12
        
        return current_clmn

    def copy_last_column(self):
        col = self.get_last_filled_column()
        self.ws.insert_cols(col + 1)

        for row in range(1, self.ws.max_row + 1):
            source_cell = self.ws.cell(row=row, column=col)
            target_cell = self.ws.cell(row=row, column=col + 1)

            # copy value or formula
            target_cell.value = source_cell.value

            # copy style
            if source_cell.has_style:
                target_cell._style = copy.copy(source_cell._style)

        return col + 1