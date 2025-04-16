import os
import pdfminer
import pdfminer.high_level
import pandas as pd
import numpy as np

# Ensure 'done' folder exists
if not os.path.exists('done'):
    os.makedirs('done')

# Process PDF files: extract text and move them to 'done'
texts = []
for file in os.listdir('.'):
    if file.endswith('.pdf'):
        with open(file, 'rb') as f:
            text = pdfminer.high_level.extract_text(f)
            texts.append(text)
        os.rename(file, os.path.join('done', file))

# Split extracted text into pages using the page break character
pages = []
for text in texts:
    pages.extend(text.split('\x0c'))

# Create DataFrame of pages and split the text on newline
df = pd.DataFrame(pages, columns=['text'])
df = df.text.str.split('\n', expand=True)

# Replace None with empty strings, then empty strings with NaN so dropna works correctly
df = df.fillna('')
df = df.replace('', np.nan)

# Drop columns and rows where all values are NaN and reset the index
df = df.dropna(axis=1, how='all').reset_index(drop=True)
df = df.dropna(axis=0, how='all').reset_index(drop=True)

# Remove illegal characters
df = df.replace(to_replace=r'\f', value='', regex=True)

# Save the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)