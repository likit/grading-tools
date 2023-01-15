import pandas as pd
from pathlib import Path
import click


@click.command()
@click.option('--ext', help='file extension', default='xlsx')
@click.option('--sheet', help='sheet name', default='problem')
@click.option('--path', help='file path', default='.')
def main(sheet, path, ext):
    p = Path(path)
    for f in p.glob(f'*.{ext}'):
        print(f'Reading {f}...')
        try:
            df = pd.read_excel(f, sheet_name=sheet)
            df.to_excel(f'{f.stem}-ans{f.suffix}', index=False)
        except Exception as e:
            print(f'Cannot read from {f} => {e}.')
        else:
            print('Finished.')


if __name__ == '__main__':
    main()