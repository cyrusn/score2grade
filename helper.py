from openpyxl.cell import Cell


def bail(msg: str) -> None:
    print(msg)
    exit(1)


def isHeaderRow(cell: Cell) -> bool:
    return cell.row == 1
