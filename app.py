import streamlit as st
import pandas as pd
import io

def load_csv(file):
    if file is not None:
        return pd.read_csv(file)
    return None

def compare_files(df1, df2, key_column):
    # Perform the merge (similar to VLOOKUP)
    merged_df = pd.merge(df1, df2, on=key_column, how='outer', indicator=True)

    # Rows only in file1
    only_in_file1 = merged_df[merged_df['_merge'] == 'left_only']
    # Rows only in file2
    only_in_file2 = merged_df[merged_df['_merge'] == 'right_only']
    # Rows in both files
    in_both = merged_df[merged_df['_merge'] == 'both']

    return only_in_file1, only_in_file2, in_both

def save_to_excel(file1_only, file2_only, both):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        file1_only.to_excel(writer, sheet_name='Only in File 1', index=False)
        file2_only.to_excel(writer, sheet_name='Only in File 2', index=False)
        both.to_excel(writer, sheet_name='In Both Files', index=False)
    output.seek(0)
    return output

def main():
    st.title('CSV File Comparison Tool')

    file1 = st.file_uploader("Upload the first CSV file", type="csv")
    file2 = st.file_uploader("Upload the second CSV file", type="csv")

    if file1 and file2:
        df1 = load_csv(file1)
        df2 = load_csv(file2)

        if df1 is not None and df2 is not None:
            key_column = st.text_input("Enter the key column for comparison")

            if key_column and key_column in df1.columns and key_column in df2.columns:
                only_in_file1, only_in_file2, in_both = compare_files(df1, df2, key_column)
                
                st.write(f"Rows only in the first file:")
                st.dataframe(only_in_file1)
                
                st.write(f"Rows only in the second file:")
                st.dataframe(only_in_file2)
                
                st.write(f"Rows in both files:")
                st.dataframe(in_both)

                excel_file = save_to_excel(only_in_file1, only_in_file2, in_both)
                
                st.download_button(
                    label="Download Comparison Results",
                    data=excel_file,
                    file_name="comparison_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Key column must be present in both files.")
        else:
            st.error("Please upload valid CSV files.")
    else:
        st.info("Please upload both CSV files.")

if __name__ == "__main__":
    main()
