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
            sheet = wb['problems']
        except Exception as e:
            print(f'{f} => {e}')
            continue
        score = 0
        try:
            # Question 1
            if sheet['E4'].value == 117 and sheet['F4'].value == 203:
                score += 5
            # Question 2
            if sheet['F7'].value == 93 and sheet['F8'].value == 162:
                score += 1.5
            if sheet['G7'].value == 27 and sheet['G8'].value == 38:
                score += 1.5
            # Question 3
            if sheet['F10'].value == 30 or sheet['F10'].value == 142:
                score += 2
            # Question 4
            if not isinstance(sheet['F12'].value, str) and sheet['F12'] is not None:
                if round(sheet['F12'].value, 2) == 5.25:
                    score += 1
            if not isinstance(sheet['F13'].value, str) and sheet['F13'] is not None:
                if round(sheet['F13'].value, 2) == 5.05:
                    score += 1
            if not isinstance(sheet['F14'].value, str) and sheet['F14'] is not None:
                if round(sheet['F14'].value, 2) == 3.79:
                    score += 1
            if not isinstance(sheet['F15'].value, str) and sheet['F15'] is not None:
                if round(sheet['F15'].value, 2) == 6.50:
                    score += 1
            if not isinstance(sheet['F16'].value, str) and sheet['F16'] is not None:
                if round(sheet['F16'].value, 2) == 11.31:
                    score += 1
            if not isinstance(sheet['F17'].value, str) and sheet['F17'] is not None:
                if round(sheet['F17'].value, 2) == 1.88:
                    score += 1
            if sheet['F19'].value == 1:
                score += 3 
            if sheet['E22'].value == 93 and sheet['F22'] == 55:
                score += 1  
            if sheet['G25'].value == 17:
                score += 1
            if sheet['G26'].value == 98:
                score += 1
            if sheet['G27'].value == 19:
                score += 1
            if sheet['G28'].value == 13:
                score += 1
        except Exception as e:
            print(f'{f} failed {str(e)}')
        else:
            print(f'{f} score={score}')
    
    
if __name__ == '__main__':
    main()