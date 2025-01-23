import pandas as pd

class ExcelHandler:
    def __init__(self, filename):

        self.filename = filename
        self.data = pd.read_excel(filename)
    
    def get_first_row(self):
        return self.data.iloc[0].tolist()

    def get_last_row(self):
        return self.data.iloc[-1].tolist()

    def get_as_dict(self):
        return self.data.to_dict(orient='list')

    def read_column(self):
        column_number = int(input("Shomare Soton Ra Vared Konid:"))
        return self.data.iloc[:, column_number].tolist()
    
    def read_row(self):
        row_number = int(input("Shomare Satr Ra Vared Konid:"))
        return self.data.iloc[row_number].tolist()

    def add_row(self):
        num_columns = len(self.data.columns)
        row_data = list(map(int, input(f"Lotfan {num_columns} Adad Baraye Satre Jadid Ezafe Konid : ").split()))
        while len(row_data) != num_columns:
            print(f"Tedad Adad Vared Shode Bayad {num_columns} Bashad!")
            row_data = list(map(int, input(f"Lotfan {num_columns} Adad Baraye Satre Jadid Ezafe Konid : ").split()))

        new_data = pd.DataFrame([row_data], columns=self.data.columns)
        self.data = pd.concat([self.data, new_data], ignore_index=True)
        self.data.to_excel(self.filename, index=False)  
        print("Satre Jadid Ezafe Shod!")

    def add_column(self):
        num_rows = len(self.data)
        column_name = input("Lotfan Name Soton Jadid Ra Vared Konid : ")
        column_data = list(map(int, input(f"Lotfan {num_rows} Adad Baraye Soton '{column_name}' Vared Konid : ").split()))
        while len(column_data) != num_rows:
            print(f"Tedad Adad Vared Shode Bayad {num_rows} Bashad!")
            column_data = list(map(int, input(f"Lotfan {num_rows} Adad Baraye Soton '{column_name}' Vared Konid : ").split()))
        
        self.data[column_name] = column_data
        self.data.to_excel(self.filename, index=False) 
        print("Soton Jadid Ezafe Shod!")


if __name__ == "__main__":

    excel = ExcelHandler("data.xlsx") 

    print("Satre Aval :", excel.get_first_row())
    print("Satre Akhar :", excel.get_last_row())
    print("Khoroji Dictionary :", excel.get_as_dict())
    print("Khoroji Shomare Soton Delkhah :", excel.read_column())
    print("Khoroji Shomare Satr Delkhah :", excel.read_row())
    excel.add_row()
    excel.add_column()