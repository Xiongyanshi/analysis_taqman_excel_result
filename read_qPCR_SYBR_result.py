#! /usr/local/bin/python3
# _*_ coding:utf-8 _*_

import sys
try:
    import xlrd
except:
    print('need xlrd package, try \'pip3 install xlrd\' to install')


def usage():
    message = '''This is a python3 script to infer .txt result from TaqMan
real-time PCR platform.
to use it:
    $ pyhton3 ./read_realtime_pcr.py YourData1.txt YourData2.txt ...
Output:
    a .csv file at working dictory.
    '''
    print(message)


def txt_to_info(filenames):
    """
    yield name, target, ct, tm informations from given .txt file
    """

    for filename in filenames:
        with open(filename, 'r') as handle:
            found = False
            for i in handle:
                """
                isolate data in:
                [Results]
                dataline1...
                dataline2...
                dataline3...
                .
                .
                .
                [Next sheet]

                """
                if i.strip() == '[Results]':  # found Results sheet
                    found = True
                    continue
                elif found and len(i.split()) < 5:  # omit rest sheets
                    break
                if found and not i.strip().startswith('Well'):  # start
                    name = i.split('\t')[3]
                    target = i.split('\t')[4]
                    ct = float(i.split('\t')[8]) if i.split('\t')[8] != 'Undetermined' else 45.0
                    tm = float(i.split('\t')[24])
                    yield name, target, ct, tm


def xls_to_info(filenames):
    """
    yield sample_name, target_name, ct, tm information from given .xls excel
    """

    for filename in filenames:
        wb = xlrd.open_workbook(filename)
        try:
            result_sheet = wb.sheet_by_name('Results')
        except ValueError:
            print('no sheet named \'Results\' in {}, please check.'.format(
                filename))
            break
        found = False
        for row in result_sheet.get_rows():
            if row[0].value == 'Well':
                found = True
                continue
            elif found is False:
                continue
            name = row[3].value
            target = row[4].value
            ct = float(row[8].value) if row[8].value != 'Undetermined' else 45.0
            tm = float(row[24].value)
            yield name, target, ct, tm


def info_to_dict(filenames):
    """
    Input:
        Tuple:
            (name, target, ct, tm)
    output:
        Dict:
            {name:
                    {target:[[ct1, ct2...], [tm1, tm2...]]}
            }
    """
    res_dict = {}

    if filenames[0].endswith('.xls'):
        info_generator = xls_to_info(filenames)
    elif filenames[0].endswith('.txt'):
        info_generator = txt_to_info(filenames)
    else:
        print('Error, file format not supported!')

    for i_name, i_target, i_ct, i_tm in info_generator:
        if i_name not in res_dict:
            res_dict[i_name] = {i_target: [[i_ct, ], [i_tm, ]]}
        elif i_name in res_dict:
            if i_target in res_dict[i_name]:
                res_dict[i_name][i_target][0].append(i_ct)
                res_dict[i_name][i_target][1].append(i_tm)
            else:
                res_dict[i_name][i_target] = [[i_ct, ], [i_tm, ]]
    return res_dict


def mean(l):
    l = [float(i) for i in l]
    return sum(l) / len(l)


def makejudge(dic):
    """
    input result dictionary from getTm()
    """
    head_printed = False
    for sample, parameters in dic.items():
        target_str = ''
        result_str = ''
        count = 0.0
        success = 0.0
        if sample == '2ng gDNA':
            for target, [cts, tms] in sorted(parameters.items()):
                target_str += target
                for ct in cts:
                    if 16 < float(ct) < 30:
                        result_str += '+'
                        success += 0.5
                    else:
                        result_str += '-'
                target_str += ','
                result_str += ','
                count += 1
        else:
            for target, [cts, tms] in sorted(parameters.items()):
                target_str += target
                for tm in tms:
                    tm = float(tm)
                    if abs(tm - mean(dic['2ng gDNA'][target][1])) < 0.5:
                        result_str += '+'
                    else:
                        result_str += '-'
                if result_str.endswith('++'):
                    success += 1
                target_str += ','
                result_str += ','
                count += 1
        target_str = ',,,' + target_str
        percent = round(100 * int(success) / int(count), 2)
        result_str = '{},{}/{},{}%,{}'.format(sample,
                                              int(success), int(count),
                                              percent,
                                              result_str)
        if not head_printed:
            yield target_str
            last_target_str = target_str
            head_printed = True
        elif target_str != last_target_str:
            yield target_str
            last_target_str = target_str
        yield result_str


def main():
    filenames = sys.argv[1:]
    savefilename = filenames[0].split()[0] + '.csv'
    with open(savefilename, 'w') as handle:
        for line in makejudge(info_to_dict(filenames)):
            handle.writelines(line + '\n')
            print(line.replace(',', '\t'))


if __name__ == "__main__":
    print()
    if len(sys.argv) == 1:
        usage()
    else:
        main()
    print('\ndone!\n')
