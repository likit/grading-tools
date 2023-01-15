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
            sheet = wb['Problems']
        except Exception as e:
            print(f'{f} => {e}')
            continue
        score = 0
        try:
            if not isinstance(sheet['F7'].value, str) and sheet['F7'].value:
                if round(sheet['F7'].value, 2) == 21.42:
                    score += 1
            if not isinstance(sheet['F8'].value, str) and sheet['F8'].value:
                if round(sheet['F8'].value, 2) == 20.68:
                    score += 1
            if sheet['F12'].value == 10\
                and sheet['G11'].value == 3 and sheet['G12'].value == 24\
                and sheet['H11'].value == 1 and sheet['H12'].value == 8\
                and sheet['I11'].value == 4 and sheet['I12'].value == 1:
                score += 5
            try:
                if sheet['F16'].value == 0.3 and sheet['F17'].value == 0.0\
                    and sheet['F18'].value == 0.1 and sheet['F19'].value == 0.6\
                    and round(sheet['H16'].value, 2) == 0.11 and round(sheet['H17'].value, 2) == 0.11\
                    and sheet['H18'].value == 0.0 and round(sheet['H19'].value, 2) == 0.78\
                    and sheet['I16'].value == 0.0 and sheet['I17'].value == 0.8\
                    and sheet['I18'].value == 0.0 and sheet['I19'].value == 0.2\
                    and round(sheet['G16'].value, 2) == 0.11 and round(sheet['G17'].value, 2) == 0.11\
                    and sheet['G18'].value == 0.0 and round(sheet['G19'].value, 2) == 0.78:
                    score += 5
            except Exception:
                pass
            if sheet['F21'].value == 'แอโรบิค':
                score += 1
            if sheet['F26'].value == 7\
                and sheet['F28'].value == 4\
                and sheet['F30'].value == 8\
                and sheet['F32'].value == 8\
                and sheet['G26'].value == 10\
                and sheet['G28'].value == 17\
                and sheet['G30'].value == 22\
                and sheet['G32'].value == 15:
                score += 7

            if sheet['E35'].value and round(sheet['E35'].value, 2) == 0.97:
                    score += 2
            if sheet['E36'].value and round(sheet['E36'].value, 2) == 0.07:
                    score += 2
            if sheet['E37'].value and round(sheet['E37'].value, 2) == 0.03:
                    score += 2
            
        except Exception as e:
            print(f'{f} failed {str(e)}')
        else:
            print(f'{f} score={score}')
    
    
if __name__ == '__main__':
    main()