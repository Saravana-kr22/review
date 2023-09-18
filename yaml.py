import ruamel.yaml as yaml
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

def extract_verification_from_yaml(file_path):
    try:
        with open(file_path, 'r') as yaml_file:
            data = yaml.safe_load(yaml_file)
            if 'tests' in data and isinstance(data['tests'], list):
                verifications = [test.get('verification', '') for test in data['tests']]
                return verifications
            else:
                print("YAML data structure does not match expectations.")
                return []

    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except yaml.YAMLError as e:
        print(f"YAML parsing error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")


file_path = "test.yaml"  # Replace with the path to your YAML file
v = extract_verification_from_yaml(file_path)

file_path = 'test-plan.xlsx'  # Replace with the path to your Excel file
#sheet_name1 = 'Sheet1'       # Replace with the name of your sheet
sheet_name2 = 'Sheet2'
#tp = read_excel_to_dict(file_path, sheet_name1)
vs = read_excel_to_dict(file_path, sheet_name2)
a = 0
b = 0
#print (tp)
#print (vs)

tp = []

#print(vs)

for u in range(len(vs)):
    k = vs[u]
    #print(k)
    p = k[0]
    #print(p)
    tp.append(p)

#print(tp)


print(len(tp))

# Print the resulting dictionary
for i in range(len(tp)):
    if tp[i]==v[i]:
        a=a+1
    else:
            print(f"mismatch in {i+1}")
            b = b+1
            #print(tp[i])
            #print(v[i])

            s1 = str(tp[i])
            s2 = str(v[i])
            d = find_sentence_difference(s1 , s2)
            print (d)

print (b)


'''if v:
        print("Verification values:")
        for i in range(len(v)):
            print(v[i])
else:
        print("No verification values found.")'''