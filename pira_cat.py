# coding: utf-8
from astropy.table import Table, vstack
from pandas import read_excel

def capitalize(text, exceptions=None):
    """
    Capitalizes every word in a string except for specified exceptions.
    Ensures the first word is always capitalized, regardless of exceptions.
    """
    if exceptions is None:
        # Common exceptions (articles, short prepositions, conjunctions)
        exceptions = {'a', 'an', 'the', 'and', 'but', 'or', 'for', 'nor', 'on', 'at', 'to', 'from', 'by', 'of', 'in', 'with'}

    # Split the input text into words and convert to lowercase for consistent comparison
    word_list = text.lower().split()
    capitalized_words = []

    for i, word in enumerate(word_list):
        # Always capitalize the first word of the sentence
        if i == 0:
            capitalized_words.append(word.capitalize())
        # Capitalize the word if it is not in the exceptions set
        elif word not in exceptions:
            capitalized_words.append(word.capitalize())
        # Otherwise, keep it in lowercase
        else:
            capitalized_words.append(word)

    # Join the processed words back into a single string
    return ' '.join(capitalized_words)

piracat = None

for subj in ['Mechanics', 'Fluids', 'Oscillations', 'Thermodynamics', 'E&M', 'Optics', 'Modern', 'Astronomy']:
    cat = Table.from_pandas(read_excel('PIRA BIB 7-1-2024.xlsx', sheet_name=subj, skiprows=4, usecols='A:D', names=['Refs', 'Name', 'PIRA', 'Desc']))
    try:
        select = ['F&A' in f'{x}' for x in cat['Refs']]
        cat = cat[select]
        print(subj, len(cat))
        cat.sort('Refs')
        cat['F&A'] = [x.split(', ')[1] for x in cat['Refs']]
        cat['Category'] = subj
        cat['Name'] = [capitalize(x) for x in cat['Name']]

        if piracat is None:
            piracat = cat['Category', 'F&A', 'PIRA', 'Name', 'Desc']
        else:
            piracat = vstack([piracat, cat['Category', 'F&A', 'PIRA', 'Name', 'Desc']])
    except:
        print(cat['Refs'])

piracat.write('pira_catalog.ecsv', format='ascii.ecsv', overwrite=True)
