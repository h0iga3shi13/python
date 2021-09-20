import xlrd
import xlsxwriter
import mojimoji

def zen2han(val):
    if type(val) is not str:
        val = str(val)
    val = mojimoji.zen_to_han(val, kana=False)
    return val


def main():
    orig_book = xlrd.open_workbook('before.xls')
    orig_sheet = orig_book.sheet_by_index(0)
    write_book = xlsxwriter.Workbook('after.xls')
    write_sheet = write_book.add_worksheet()

    for row in range(orig_sheet.nrows):
        for col in range(orig_sheet.ncols):
            cell = zen2han(orig_sheet.cell(row,col).value)
            write_sheet.write(row,col,cell)
    write_book.close()


if __name__ == '__main__':
    main()