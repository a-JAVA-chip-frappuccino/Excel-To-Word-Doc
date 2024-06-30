import pandas as pd

def open_excel(input_path, output_path):
    # read excel file into dataframe
    df = pd.read_excel(file_path)
    process_df(df, output_path)

def process_df(df, output_path):
    info = ""
    for idx, row in df.iterrows():
        # creates human-readable dictionary from row of dataframe
        info += str(row.to_dict())
        if idx != len(df) - 1:
            info += "\n\n"
    write_to_doc(info, output_path)

def write_to_doc(info, output_path):
    # save info to a word document
    with open(output_path, 'w') as f:
        f.write(info)
    print(f"Document saved to {output_path}")

if __name__ == "__main__":
    file_path = ''' insert Excel file path here '''
    output_path = ''' insert Word Doc file path here '''

    open_excel(file_path, output_path)