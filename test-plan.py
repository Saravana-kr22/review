import openpyxl
import pandas as pd
import difflib



def find_sentence_difference(sentence1, sentence2):
    # Create a Differ object
    differ = difflib.Differ()
    
    if isinstance(sentence1, str):
 
        # Compute the difference between the two sentences
        diff = list(differ.compare(sentence1.split(), sentence2.split()))
        
        # Join the differences into a single string
        diff_string = ' '.join(diff)
        
        return diff_string
    else:
        return "none"


def read_excel_to_dict(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Convert the DataFrame to a dictionary
    data_dict = df.to_dict(orient='split')

    # Extract the data from the dictionary
    data = {}
    for key, value in zip(data_dict['index'], data_dict['data']):
        data[key] = value

    return data

# Example usage:
file_path = 'test-plan.xlsx'  # Replace with the path to your Excel file
sheet_name1 = 'Sheet1'       # Replace with the name of your sheet
sheet_name2 = 'Sheet2'
tp = read_excel_to_dict(file_path, sheet_name1)
vs = read_excel_to_dict(file_path, sheet_name2)
a = 0
b = 0
#print (tp)
#print (vs)

c = ['#', 'Ref', 'PICS', 'Test Step' , 'Expected Outcome']

# Print the resulting dictionary
for i in range(len(tp)):
    l1 = tp[i]
    l2 = vs[i]
    #print(tp[i])
    for j in range(len(l1)):
        if l1[j] == l2[j]:
            a = a+1

        else:
            print(f"mismatch in {c[j]} on {l1[0]}")
            b = b+1

            print(l1[j])
            print(l2[j])
            s1 = str(l1[j])
            s2 = str(l2[j])
            d = find_sentence_difference(s1 , s2)
            print (d)
print (b)
