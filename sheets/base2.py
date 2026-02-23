import copy

class BaseSheet2:
    """
    Base class for Dashboard sheet processing.
    Handles deletion of first column (C) and shifting all data forward,
    then copying the last column to create a new month column.
    """
    
    def __init__(self, ws, sheet_config, input_data):
        self.ws = ws
        self.cfg = sheet_config
        self.input = input_data

    def delete_and_shift_columns(self, start_row, end_row, year_row, start_col=3, end_col=14):
        """
        Deletes column C (index 3) for the specified row range and shifts columns D-N forward.
        Then copies the last column to position N.
        SKIPS the year row because it contains merged cells.
        
        Args:
            start_row: First row of the table section
            end_row: Last row of the table section
            year_row: Row containing the year (will be skipped during shifting)
            start_col: Starting column (default 3 = C)
            end_col: Ending column (default 14 = N)
        
        Returns:
            The column index where new data should be written (N = 14)
        """
        
        # Step 1: Shift all columns D to N one position to the left (into D to M)
        # SKIP the year row because it contains merged cells
        for row in range(start_row, end_row + 1):
            if row == year_row:
                continue  # Skip year row - will be handled separately
                
            for col in range(start_col, end_col):
                source_cell = self.ws.cell(row=row, column=col + 1)
                target_cell = self.ws.cell(row=row, column=col)
                
                # Copy value or formula
                target_cell.value = source_cell.value
                
                # Copy style
                if source_cell.has_style:
                    target_cell._style = copy.copy(source_cell._style)
        
        # Step 2: Copy column M (now containing former N data) to column N
        # Again SKIP the year row
        for row in range(start_row, end_row + 1):
            if row == year_row:
                continue  # Skip year row
                
            source_cell = self.ws.cell(row=row, column=end_col - 1)  # M (13)
            target_cell = self.ws.cell(row=row, column=end_col)       # N (14)
            
            # Copy value or formula
            target_cell.value = source_cell.value
            
            # Copy style
            if source_cell.has_style:
                target_cell._style = copy.copy(source_cell._style)
        
        return end_col  # Return column N (14)
