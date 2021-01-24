import openpyxl
from pathlib import Path
import networkx as nx
import matplotlib.pyplot as plt


def process(sheet_obj, sheet_name):
    G = nx.Graph()

    # Will print a particular row value

    for j in range(2, sheet_obj.max_row + 1):
        source = sheet_obj.cell(row=j, column=1)
        target = sheet_obj.cell(row=j, column=2)

        if source.value:
            G.add_node(source.value)

        if target.value:
            G.add_node(target.value)

        if source.value and target.value and \
                source.value != target.value and \
                not G.has_edge(source.value, target.value) and \
                not G.has_edge(target.value, source.value):
            G.add_edge(source.value, target.value)

        # print(row)
    plt.figure(3, figsize=(50, 50))
    nx.draw_networkx(G, arrowsize=6, with_labels=True, node_size=10, font_size=5)
    plt.savefig(sheet_name + ".pdf")


def main():
    loc = 'RED SDH.xlsx'
    wb_obj = openpyxl.load_workbook(loc)
    print(wb_obj.sheetnames)

    for sheet_name in wb_obj.sheetnames:
        process(wb_obj[sheet_name], sheet_name)


if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
