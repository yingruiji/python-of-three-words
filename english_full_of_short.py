import pandas as pd

excelIn = 'test1.xlsx'
excelOut = 'testout.xlsx'
sheets = []
n = 0

sheetsName = list(pd.read_excel(excelIn, sheet_name=None))

for name in sheetsName:
    print(name)
    df = pd.read_excel(excelIn, sheet_name=name)
    sheets.append(df)

colFullWord = '复制全称'
colShorten = '英文缩写'

for sheet in sheets:
    for index, row in sheet.iterrows():
        wordsNoDot = row[colFullWord].replace('.', '')
        words = wordsNoDot.split(' ')
        if words[0] == 'The':
            words.pop(0)
        for i in range(len(words)):
            if i == 0:
                words[i] = words[i][:3].lower()
            else:
                words[i] = words[i][:3].capitalize()
        wordsOut = ''.join(words)
        sheet.loc[index, colShorten] = wordsOut
        print(index, row[colFullWord], wordsOut, row[colShorten])

excelWriter = pd.ExcelWriter(excelOut)

for n in range(len(sheets)):
    sheets[n].to_excel(excelWriter, sheet_name=sheetsName[n], index=False)
excelWriter.save()
