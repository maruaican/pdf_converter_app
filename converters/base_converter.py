import os

class BaseConverter:
    def __init__(self, file_path):
        self.file_path = os.path.abspath(file_path)
    
    def convert(self):
        raise NotImplementedError