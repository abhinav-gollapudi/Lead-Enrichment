import os
import uuid
import streamlit as st
from io import BytesIO
import os
from openai import OpenAI
import base64
import pandas as pd
import re
import json
from concurrent.futures import ThreadPoolExecutor, as_completed

load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
openai = OpenAI(api_key=OPENAI_API_KEY)

def remove_duplicates(df):
    df['num_blanks'] = df.apply(lambda row: row.isna().sum() + (row == '').sum(), axis=1)

    # Step 2: Sort by name and number of blanks (ascending to keep fewer blanks)
    df_sorted = df.sort_values(by=['Name', 'num_blanks'])

    # print(df_sorted)

    # Step 3: Drop duplicates based on the 'name' column, keeping the first (fewer blanks) row
    df_unique = df_sorted.drop_duplicates(subset='Name', keep='first').drop(columns='num_blanks')

    
    return df_unique

def parse_json(data_text):
    """
    Parse a json from plain text into a list of dictionaries.
    
    Args:
        data_text (str): The text containing the json to parse
    
    Returns:
        list: A list of dictionaries containing parsed data
    """
    # Extract lines from the json
    cleaned_text = re.sub(r'^```json\s*|\s*```$', '', data_text.strip())

    # Parse the cleaned JSON text
    parsed_list = json.loads(cleaned_text)

    # Print the parsed list
    return(parsed_list)


def encode_image(image_file):
    return base64.b64encode(image_file.read()).decode('utf-8')

def get_response(image_file):
  
  base64_image = encode_image(image_file)

  response = openai.chat.completions.create(
      model="gpt-4o",
      temperature=0,
    messages=[
      {
        "role": "user", 
        "content": [
        {"type": "text", "text": "Analyze the given image to accurately extract the full name, job title, and associated company of every individual profile displayed. Make sure to include all profiles in the image. Output the data as a JSON object with each object containing 'Name,' 'Job Title,' and 'Company,' fields. Ignore all irrelevant text, icons, or visual elements."},
          {
            "type": "image_url", 
            "image_url": {
              "url": f"data:image/png;base64,{base64_image}"
            }
          }
        ]
      }
    ]
  )
  return response.choices[0].message.content


def get_data(uploaded_files):
    """
    Concatenate multiple uploaded files into a single file.
    
    Args:
        uploaded_files (list): List of uploaded file objects
    
    Returns:
        bytes: Concatenated file content
    """
    with ThreadPoolExecutor(max_workers=min(len(uploaded_files), 10)) as executor:
        # Create a dictionary to map futures to their original files
        parsed_data = []
    
    # Use ThreadPoolExecutor for concurrent image processing
    with ThreadPoolExecutor(max_workers=min(10, len(uploaded_files))) as executor:
        # Create a dictionary to map futures to their original files
        future_to_file = {
            executor.submit(get_response, file): file 
            for file in uploaded_files
        }
        
        for future in as_completed(future_to_file):
            try:
                output = future.result()
                parsed_data.extend(parse_json(output))
            except Exception as exc:
                st.error(f'Image processing generated an exception: {exc}')
    
    return parsed_data

def main():
    # Set page configuration
    st.set_page_config(
        page_title="File Concatenator", 
        page_icon=":paperclip:",
        layout="centered"
    )
    
    # App title and description
    st.title("📄 Lead Enrichment Tool")
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose files to merge", 
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        help="Select one or more files to merge"
    )
    
    # Concatenation and download process
    if uploaded_files:
    
        # Concatenation button
        if st.button("Merge Files", type="primary"):
            try:
                with st.spinner("Processing files, please wait..."):
                    # Perform concatenation
                    parsed_data = get_data(uploaded_files)
                    
                    if parsed_data:
                        # Generate unique filename
                        df = pd.DataFrame(parsed_data)
                        df=remove_duplicates(df)

                        output_filename = f"merged_{uuid.uuid4()}.xlsx"

                        output = BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            # Write the dataframe to the Excel sheet
                            df.to_excel(writer, index=False, sheet_name="Extracted Data")
                            
                            # Get the xlsxwriter workbook and worksheet objects
                            workbook = writer.book
                            worksheet = writer.sheets["Extracted Data"]
                            
                            # Adjust column widths based on the longest content in each column
                            for i, col in enumerate(df.columns):
                                # Find the maximum length of content in the column
                                max_len = max(
                                    df[col].astype(str).map(len).max(),  # Longest content in the column
                                    len(col)  # Length of the column header
                                )
                                
                                # Set column width (add a little extra padding)
                                worksheet.set_column(i, i, max_len + 2)
                        
                        output.seek(0)

                        # Step 5: Download button for the Excel file
                        st.download_button(
                            label="Download Extracted Data",
                            data=output,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # Success message
                        st.success(f"Files merged successfully! Filename: {output_filename}")
            
            except Exception as e:
                st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
