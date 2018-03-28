#! /usr/local/bin/python3
# _*_ coding:utf-8 _*_


import sys
import xlrd
import xlwt


def usage():
    message = """This is a python script to extract 'Results' sheet data
    exported from TaqMan realtime PCR machine.

    module needed: xlrd, xlwt

    To execute:
        python3 read_qPCR_TaqMan_result.py data.xls(s)
    """
    print(message)


def xls2resultrow(fnames):
    """
    get ct information from given excel file(s)
    output:
        (samplename(str), target(str), ct(float OR str))
    """
    for fname in fnames:
        try:
            wb = xlrd.open_workbook(fname)
        except:
            print('error when loading excel file')
            raise

        result_sheet = wb.sheet_by_name('Results')

        found = False
        for row in result_sheet.get_rows():
            if row[0].value == 'Well':
                found = True
                yield row
                continue
            if found and len(row) > 1:
                yield row


def main():
    if len(sys.argv) == 1:
        return usage()
    filenames = sys.argv[1:]
    savefilename = 'result.xls'
    wbk = xlwt.Workbook(encoding='utf-8')
    sheet = wbk.add_sheet('results')

    for row_idx, row_line in enumerate(xls2resultrow(filenames)):
        for column_idx, cell in enumerate(row_line):
            sheet.write(row_idx, column_idx, cell.value)

    wbk.save(savefilename)
    print('data saved in file: {}'.format(savefilename))


if __name__ == '__main__':
    main()
    print('done!')
