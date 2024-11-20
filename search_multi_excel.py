import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")

def search_multi_excel(search_term,directory):
    results={}
    not_found_files=[]

    for file_name in [f for f in os.listdir(directory) if f.endswith('.xlsx')]:
        file_path = os.path.join(directory,file_name)
        found = False

        try:
            for sheet_name, df in pd.read_excel(file_path,sheet_name=None).items():
                if df.applymap(lambda cell: search_term.lower() in str(cell).lower()).any().any():
                    found = True
                    results.setdefault(file_name,[]).append(sheet_name)
        except Exception as e:
            print(f"Error processing {file_path}: {e}")

        if not found:
            not_found_files.append(file_name)

    return results, not_found_files
    
def main():
    directory = os.path.dirname(os.path.abspath(__file__))

    while (search_term := input("Enter CB number (e.g., CB-104259.MOSS) or 'exit' to quit: ").strip()) != 'exit':
        results, not_found_files = search_multi_excel(search_term,directory)
        print()
        if results:
            print("Found in following showrooms:")
            print("\n".join(f" - {file}" for file in results))
            print()
        if not_found_files:
            print("Not found in following showrooms:")
            print("\n".join(f" - {file}" for file in not_found_files))
            print()
        
if __name__ == "__main__":
    main()