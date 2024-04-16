import pandas as pd

dic={'cat':['a','b','c']}

a=pd.DataFrame(dic)

print(a)

def read_input_from_vba(file_path):
    with open(file_path, 'r') as file:
        input_data = file.read()
    return input_data

# Path to the file where VBA writes input
file_path = r"D:\precsion_pro\VBA_Python_connection\input.txt"

# Read input from VBA
input_from_vba = read_input_from_vba(file_path)
print("Input from VBA:", input_from_vba)


# a.to_csv(r'D:\precsion_pro\VBA_Python_connection\output\done_1.csv')



