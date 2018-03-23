import pandas as pd
import matplotlib.pyplot as plt

colonnes_a_filtrer = [
    'Annee', 'Source audience', 'Saisonnalite', 'Code Reseau', 'Source', 'Source diffusion', 'Commercialisation', 'Telephone', 'Mail '
]
fichier = 'Export_Top_Affinite.xls'


def get_file(file_path):
    xlsx = pd.ExcelFile(file_path)
    sheets = []
    for sheet in xlsx.sheet_names:
        sheets.append(xlsx.parse(sheet))
    return pd.concat(sheets)


def filter_data(sheet):
    a = sheet.drop(columns=colonnes_a_filtrer, axis=0)
    return a


def graphic(sheet):
    table = sheet.pivot_table(index='Utilite')
    ax = plt.subplot(111, frame_on=False)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_visible(False)
    table.plot(kind='bar', figsize=(20,8))
    plt.savefig('graphique.png')


def write_as_exel(sheet):
    sheet.to_excel('export_python.xlsx', index=False)


if __name__ == '__main__':
    data = get_file(fichier)
    data = filter_data(data)
    graphic(data)
    write_as_exel(data)
