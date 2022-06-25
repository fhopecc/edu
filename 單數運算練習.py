from pathlib import Path
wdir = Path(__file__).parent
wordexe = 'C:\Program Files\Microsoft Office\Office14\WINWORD.EXE'
def 單數運算評量(運算='X', 列印=False):
    運算名稱 = {'X':'乘法', '+':'加法'}[運算]
    table = [(x,y) for x in range(1,10) for y in range(1,10)]
    from random import shuffle
    shuffle(table)
    from docx import Document
    f = wdir / '模版.docx'
    doc = Document(f)
    doc.paragraphs[0].text = f"九九{運算名稱}評量"
    doc.paragraphs[0].style = doc.styles['Heading 1']
    tab = doc.add_table(rows=9, cols=9)
    for i, c in enumerate(table):
        n1, n2 = c
        c = tab.cell(i%9, i//9)
        if 運算=='X':
            c.text = (f'{n1}X{n2}=(    )')
        elif 運算=='+': 
            c.text = (f'{n1}+{n2}=(    )')
        else:
            raise ValueError('不支援【{運算}】運算題目！')
    doc.add_paragraph()
    doc.add_paragraph()
    result = '答錯題數：____題；答題時間：____分鐘____秒；評量日期：__年__月__日；姓名：________'
    doc.add_paragraph(result)
    doc.add_page_break()
    doc.add_paragraph('答案頁', style='Heading 1')
    tab = doc.add_table(rows=9, cols=9)
    for i, c in enumerate(table):
        n1, n2 = c
        c = tab.cell(i%9, i//9)
        if 運算=='X':
            c.text = (f'{n1}X{n2}=({n1*n2})')
        elif 運算=='+': 
            c.text = (f'{n1}+{n2}=({n1+n2})')
 
    fn = wdir / f'九九{運算名稱}評量.docx'
    doc.save(fn)
    from os import system
    cmd = f'start {fn}'
    if 列印:
        cmd = f'"{wordexe}" {fn} /mFilePrintDefault /mFileExit /q /n'
    system(cmd)

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("operator", help="運算類型")
    parser.add_argument("--print", action='store_true', help="列印")
    args = parser.parse_args()
    if args.print:
        單數運算評量(args.operator, 列印=True)
    else:
        單數運算評量(args.operator)
