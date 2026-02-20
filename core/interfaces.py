from abc import ABC, abstractmethod

class SheetProcessor(ABC):
    def __init__(self, workbook, config, input_data):
        self.workbook = workbook
        self.config = config
        self.input = input_data

    @abstractmethod
    def process(self):
        pass