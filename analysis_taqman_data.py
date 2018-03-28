#! /usr/local/bin/python3
# _*_ coding:utf-8 _*_


import sys
import xlrd
import xlwt
import time


def usage():
    message = """to analysis taqman realtime pcr exported excel result
Input:
    excel files
Output:
    a new excel with ct values for each targets of each samples
To execute:
    python3 analysis_taqman_data.py ./filename.excel(s)
"""
    print(message)


def xls2stc(fnames):
    """
    get ct information from given excel file(s)
    input:
        excel file name(s)
    output:
        (Samplename<str>, Target<str>, Ct<float OR str>)
    """
    for fname in fnames:
        wb = xlrd.open_workbook(fname)
        result_sheet = wb.sheet_by_name('Results')

        found = False
        for row in result_sheet.get_rows():
            if row[0].value == 'Well':
                for col_idx, col_item in enumerate(row):
                    col_value = col_item.value
                    if col_value == "Sample Name":
                        sample_name_idx = col_idx
                        continue
                    if col_value == 'Target Name':
                        target_name_idx = col_idx
                        continue
                    if col_value == "CT":
                        Ct_idx = col_idx
                        continue
                found = True
                continue
            if found and len(row) > 1:
                sample = row[sample_name_idx].value
                target = row[target_name_idx].value
                ct = row[Ct_idx].value
                if isinstance(ct, float):
                    ct = round(ct, 3)
                else:
                    ct = 'Und'
                yield (sample, target, ct)


def unpack(a_list):
    """
    unpack a list, return a new flat list
    """
    res = []
    for i in a_list:
        if isinstance(i, list):
            res += unpack(i)
        else:
            res.append(i)
    return res


def main():
    if len(sys.argv) == 1:
        return usage()
    filenames = sys.argv[1:]
    res_dict = {}
    for samplename, target, ct in xls2stc(filenames):
        if samplename not in res_dict:
            res_dict[samplename] = {target: [ct, ]}
            continue
        elif target not in res_dict[samplename]:
            res_dict[samplename][target] = [ct, ]
            continue
        else:
            res_dict[samplename][target].append(ct)
    """
    for sample in res_dict:
        print(sample, '\t', res_dict[sample])
    """

    # prepare a excel file and a 'results' sheet.
    savefilename = 'result.xls'
    wbk = xlwt.Workbook(encoding='utf-8')
    sheet = wbk.add_sheet('results')

    # going to write rows
    head_printed = False
    row_idx = 0
    for sample in res_dict:
        target_list = sorted(res_dict[sample])
        ct_list = [res_dict[sample][target] for target in target_list]
        target_list = [[i]*len(j) for i, j in zip(target_list, ct_list)]

        # some sample cases have too much Ct values,like PC or NC.
        if len(ct_list[0]) > 5:
            continue

        # first time come to write, need to write head
        if not head_printed:
            for column_idx, cell in enumerate(unpack(target_list)):
                sheet.write(row_idx, column_idx+1, cell)
            print(target_list)
            row_idx += 1
            head_printed = True
            last_target_list = target_list

        # if target list changed, re-write head(target list)
        if last_target_list != target_list:
            for column_idx, cell in enumerate(unpack(target_list)):
                sheet.write(row_idx, column_idx+1, cell)
            last_target_list = target_list
            print(target_list)
            row_idx += 1

        # then write ct values
        row_items = [sample] + unpack(ct_list)
        for column_idx, cell in enumerate(row_items):
            sheet.write(row_idx, column_idx, cell)
        print(row_items)
        row_idx += 1

    wbk.save(savefilename)
    print('data saved in file: {}'.format(savefilename))


if __name__ == '__main__':
    main()
    print('done!')
