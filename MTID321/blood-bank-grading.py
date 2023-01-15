from pathlib import Path
import click
from openpyxl import load_workbook


@click.command()
@click.option('--ext', help='file extension', default='xlsx')
@click.option('--path', help='file path', default='.')
def main(path, ext):
    p = Path(path)
    for f in p.glob(f'*.{ext}'):
        try:
            wb = load_workbook(f, data_only=True)
            sheet = wb['Sheet1']
        except Exception as e:
            print(f'{f} => {e}')
            continue
        score = 0
        try:
            if isinstance(sheet['E8'].value, str) and isinstance(sheet['F8'].value, str) and\
                isinstance(sheet['G8'].value, str) and isinstance(sheet['H8'].value, str):
                if 'A' in sheet['E8'].value and 'O' in sheet['F8'].value\
                    and 'O' in sheet['G8'].value and 'B' in sheet['H8'].value:
                    score += 2
            if isinstance(sheet['E11'].value, str) and isinstance(sheet['F11'].value, str) and\
                isinstance(sheet['G11'].value, str) and isinstance(sheet['H11'].value, str):
                if 'A' in sheet['E11'].value and 'O' in sheet['F11'].value\
                    and 'AB' in sheet['G11'].value and 'B' in sheet['H11'].value:
                    score += 2
            if ('SUR' in sheet['E14'].value or 'OTH' in sheet['E14'].value)\
                and sheet['F14'].value == 'PED' and sheet['G14'].value == 'SUR'\
                and sheet['H14'].value == 'GYN':
                score += 2
            if sheet['E18'].value == 'PED' and sheet['F18'].value == 'ER'\
                and sheet['G18'].value == 'MED' and sheet['H18'].value == 'OPD':
                score += 2
            if sheet['E19'].value == 21.45 and sheet['F19'].value == 28.35\
                and sheet['G19'].value == 19.00 and sheet['H19'].value == 19.19:
                score += 1
            if sheet['E23'].value == 230100 and sheet['F23'].value == 220700\
                and sheet['G23'].value == 216200:
                score += 3
            if sheet['D25'].value == 'ER':
                score += 1
            if sheet['E28'].value == 'ภัทรวดี' and sheet['F28'].value == 'สมศักดิ์':
                score += 2
            if isinstance(sheet['D31'].value, str) and isinstance(sheet['E31'].value, str) and\
                isinstance(sheet['F31'].value, str):
                if 'ER' in sheet['D31'].value and 'PED' in sheet['E31'].value\
                    and 'OTH' in sheet['F31'].value:
                    score += 3
            if sheet['C38'].value == 5:
                score += 2
        except Exception as e:
            print(f'{f} failed {str(e)}')
        else:
            print(f'{f} score={score}')
    
    
if __name__ == '__main__':
    main()