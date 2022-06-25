from pathlib import Path
wdir = Path(__file__).parent

def 單數運算練習(運算='X'):
    運算名稱 = {'X':'乘法', '+':'加法'}[運算]
    table = [(x,y) for x in range(1,10) for y in range(1,10)]
    from random import shuffle
    shuffle(table)
    from docx import Document
    f = wdir / '模版.docx'
    doc = Document(f)
    doc.paragraphs[0].text = f"__年__月__日弼叡的九九{運算名稱}練習"
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
    doc.add_paragraph('本次錯誤題數：_____題，使用秒數：____秒')
    doc.add_page_break()
    doc.add_paragraph('答案', style='Heading 1')
    tab = doc.add_table(rows=9, cols=9)
    for i, c in enumerate(table):
        n1, n2 = c
        c = tab.cell(i%9, i//9)
        if 運算=='X':
            c.text = (f'{n1}X{n2}=({n1*n2})')
        elif 運算=='+': 
            c.text = (f'{n1}+{n2}=({n1+n2})')
 
    fn = wdir / f'九九{運算名稱}練習.docx'
    doc.save(fn)
    from os import system
    system(f'start {fn}')

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("operator", help="運算")
    args = parser.parse_args()
    單數運算練習(args.operator)
