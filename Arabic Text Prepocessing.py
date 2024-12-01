import os
import re
import pandas as pd
from tashaphyne.normalize import strip_tashkeel, strip_tatweel, normalize_hamza, normalize_lamalef, normalize_spellerrors

def kashida_removal(text):
    text = re.sub(r'ـ+', '', text)
    return text

def arabic_diacritics_removal(text):
    diacritics = 'ًٌٍَُِّّْٰ'
    for d in diacritics:
        text = text.replace(d, '')
    return text

def remove_punctuation(text):
    punctuations = r"""!"#$%&'()*+,-./:;<=>?@[\]^_‘;`{|}~•«»…“”–—٪"""
    arabic_punctuations = r"؟،؛ـ"
    all_punctuations = punctuations + arabic_punctuations
    translator = str.maketrans('', '', all_punctuations)
    return text.translate(translator)

def alef_lam_normalization(text):
    text = re.sub(r'[إأآا]', 'ا', text)
    text = re.sub(r'[يى]', 'ي', text)
    text = re.sub(r'[ؤئ]', 'ء', text)
    text = re.sub(r'[ةه]', 'ه', text)
    return text

def normalize_text(text):
    new_text = strip_tashkeel(text)
    new_text = strip_tatweel(new_text)
    new_text = normalize_hamza(new_text)
    new_text = normalize_lamalef(new_text)
    new_text = arabic_diacritics_removal(new_text)
    new_text = alef_lam_normalization(new_text)
    new_text = kashida_removal(new_text)
    new_text = remove_punctuation(new_text)
    new_text = normalize_spellerrors(new_text)
    return new_text

# Specify the directory where the files are located
directory = 'ADD YOUR DIRECTORY'

# Iterate over all files in the directory
for filename in os.listdir(directory):
    filepath = os.path.join(directory, filename)

    if filename.endswith('.xlsx'):
        # If the file is an Excel file
        df = pd.read_excel(filepath, header=0)  # Assuming the first row contains column names

        # Preprocess the text in each column
        for column in df.columns:
            df[column] = df[column].apply(lambda x: normalize_text(str(x)) if pd.notna(x) else '')

        # Save the DataFrame back to the same Excel file
        df.to_excel(filepath, index=False)

    elif filename.endswith('.txt'):
        # If the file is a text file
        with open(filepath, 'r', encoding='utf-8') as file:
            content = file.readlines()

        # Normalize each line in the text file
        normalized_content = [normalize_text(line.strip()) for line in content]

        # Write the normalized content back to the text file
        with open(filepath, 'w', encoding='utf-8') as file:
            for line in normalized_content:
                file.write(f"{line}\n")
