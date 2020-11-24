# tmp env
from PyPDF2 import PdfFileReader
import pandas as pd
import os

a1_directory = 'CompactExamsAI-EV'
geo_directory = 'CompactExamsGEO-EV'
a2_directory = 'CompactExamsAII-EV'

month_dict = {
    '01': 'January',
    '06': 'June',
    '08': 'August'
}


def df_maker(directory, pdf_file):
    path = os.path.join(directory, pdf_file)
    with open(path, 'rb') as file:
        fileReader = PdfFileReader(file)
        text = ''
        for page in range(fileReader.getNumPages()):
            if fileReader.getPage(page).extractText()[:6] == ' ID: A':
                text += fileReader.getPage(page).extractText()
        topics = text.split('TOP:')
        top_list = []
        for topic in topics[1:]:
            top_list.append(topic.split('\n')[1].strip())

        standards = text.split('NAT:')
        std_list = []
        for topic in standards[1:]:
            std_list.append(topic.split('\n')[1].strip())

    df = pd.DataFrame(zip(range(1, len(top_list)+1), top_list, std_list),
                      columns=['Question', 'Topic', 'Standard'])

    return df


for dir in [a1_directory,
            geo_directory,
            a2_directory]:
    if dir == a1_directory:
        file_name = 'Algebra I.xlsx'
    elif dir == geo_directory:
        file_name = 'Geometry.xlsx'
    elif dir == a2_directory:
        file_name = 'Algebra II.xlsx'

    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

    for filename in os.listdir(dir):
        if filename[:2] in ('01', '06', '08'):
            month = month_dict[filename[:2]]
            year = '20'+filename[2:4]
            df = df_maker(dir, filename)
            df.to_excel(writer, sheet_name=f'{year} {month}', index=False)
            print(filename, "done")
    writer.save()

print("All done")
