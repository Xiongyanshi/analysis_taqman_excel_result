#! /usr/local/bin/python3
# _*_ coding:utf-8 _*_


import sys
import xlrd
import xlwt
import time


def usage():
    message = """***************
A python3 script to analysis Life Real-Time PCR Software
exported excel result.
---------------
Input:
    .xls file(s)
Output:
    a new .xls file with ct values of each targets for each samples.
To execute:
    python3 analysis_taqman_data.py ./filename.excel(s)
---------------
Author: Yanshi, Xiong
E-mail: 1240172230@qq.com
***************
"""
    print(message)


def xls2stc(fname):
    """
    get ct information from given excel file
    input:
        excel file name
    output:
        (Samplename<str>, Target<str>, Ct<float OR str>)
    """
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
    my powerful magic function, hahaha.
    """
    res = []
    for i in a_list:
        if isinstance(i, list):
            res += unpack(i)
        else:
            res.append(i)
    return res


def stc2dict(filename):
    """
    for a given .xls file, build a dict structure to hold data
    structure:
        {sample1:
                {target1:[ct1, ct2],
                 target2:[ct1, ct2]
                 ...
                 ...
                 'gender':'female'/'male'/'undetermined'
                }
         sample2:
                {target1:[ct1, ct2],
                 target2:[ct1, ct2]
                 ...
                 ...
                 'gender':'female'/'male'/'undetermined'
                }
        ...
        more samples
        ...
        }
    """
    res_dict = {}
    for sample, target, ct in xls2stc(filename):
        if sample not in res_dict:
            res_dict[sample] = {target: [ct]}
            continue
        elif target not in res_dict[sample]:
            res_dict[sample][target] = [ct]
        else:
            res_dict[sample][target].append(ct)

    """make gender judge"""
    pass
    return res_dict


def main():
    """
    get excel filename(s), read data in each file into dict
    write dict data into new excel file.
    """
    if len(sys.argv) == 1:
        return usage()
    filenames = sys.argv[1:]

    # prepare a excel file and a 'results' sheet.
    savefilename = time.asctime().replace(' ', '_') + '.xls'
    wbk = xlwt.Workbook(encoding='utf-8')
    sheet = wbk.add_sheet('results')

    # going to write rows
    # initiating...
    last_target_list = []
    row_idx = 0
    print('Start read and analysis data...')
    for filename in filenames:
        data_dict = stc2dict(filename)
        experiment_id = filename.rstrip('.xls')
        print('File: {}'.format(filename))
        for sample in data_dict:
            target_list = sorted(data_dict[sample])
            ct_list = [data_dict[sample][target] for target in target_list]
            target_list = [[i]*len(j) for i, j in zip(target_list, ct_list)]

            # if target items changed, write new table head.
            if target_list != last_target_list:
                for column_idx, cell in enumerate(unpack(target_list)):
                    # first 2 columns blanked
                    sheet.write(row_idx, column_idx+2, cell)
                last_target_list = target_list
                row_idx += 1

            # write sample experimant result each time.
            print('  sample: {}'.format(sample))

            row_items = [experiment_id, sample] + unpack(ct_list)
            for col_idx, cell in enumerate(row_items):
                sheet.write(row_idx, col_idx, cell)
            row_idx += 1

    wbk.save(savefilename)
    print('\nData saved in file: {}'.format(savefilename))


if __name__ == '__main__':
    main()
    print('done!')
