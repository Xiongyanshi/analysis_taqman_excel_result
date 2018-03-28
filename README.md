This is a python3 script to analysis TaqMan realtime PCR result

# Introduction
realtime PCR machine model: Life QuantStudio 6 Flex
running software for this equipment: QuantStudio Real-Time PCR Software(version 1.1)

## Wet-lab experimant information:
We are quite interested in some specific DNA target in our biological samples.
So, we want to quantitative analysis it. With well designed primer pairs and probe,
Taqman realtime PCR is the best choice to make it posible.

This Python3 script is the tool we use to make anonying and error-prone jobs done
by computer, especially when you have a large number of excel files.

## What happened in this Python3 script:
In each Excel file QuantStudio machine exported, the information we care most is in 'Results' sheet.
In this sheet, each line(or say Row) we interested most is three columns: 
Sample name; target name, ct value.
|      |             |       |             |       |          |      |
|------|-------------|-------|-------------|-------|----------|------| 
| ...  | Sample Name | ...   | Target Name | ...   | Ct value |      | 
| ...  |  PC         |  ...  |  target1    |  ...  |  ct1     |  ... | 
| ...  |  PC         |  ...  |  target1    |  ...  |  ct2     |  ... | 
| ...  |  PC         |  ...  |  target2    |  ...  |  ct3     |  ... | 
| ...  |  PC         |  ...  |  target2    |  ...  |  ct4     |  ... | 
| ...  |  PC         |  ...  |  target3    |  ...  |  ct5     |  ... | 
| ...  |  PC         |  ...  |  target3    |  ...  |  ct6     |  ... | 
| ...  |  sample1    |  ...  |  target1    |  ...  |  ct7     |  ... | 
| ...  |  sample1    |  ...  |  target1    |  ...  |  ct8     |  ... | 
| ...  |  sample1    |  ...  |  target2    |  ...  |  ct9     |  ... | 
| ...  |  sample1    |  ...  |  target2    |  ...  |  ct10    |  ... | 
| ...  |  sample1    |  ...  |  target3    |  ...  |  ct11    |  ... | 
| ...  |  sample1    |  ...  |  target3    |  ...  |  ct12    |  ... | 
| ...  |  sample2    |  ...  |  target1    |  ...  |  ct13    |  ... | 
| ...  |  sample2    |  ...  |  target1    |  ...  |  ct14    |  ... | 
| ...  |  sample2    |  ...  |  target2    |  ...  |  ct14    |  ... | 
| ...  |  sample2    |  ...  |  target2    |  ...  |  ct16    |  ... | 
| ...  |  sample2    |  ...  |  target3    |  ...  |  ct17    |  ... | 
| ...  |  sample2    |  ...  |  target3    |  ...  |  ct18    |  ... | 
| ...  |  sample3    |  ...  |  target1    |  ...  |  ct19    |  ... | 
| ...  |  sample3    |  ...  |  target1    |  ...  |  ct20    |  ... | 
| ...  |  sample3    |  ...  |  target2    |  ...  |  ct21    |  ... | 
| ...  |  sample3    |  ...  |  target2    |  ...  |  ct22    |  ... | 
| ...  |  sample3    |  ...  |  target3    |  ...  |  ct23    |  ... | 
| ...  |  sample3    |  ...  |  target3    |  ...  |  ct24    |  ... | 
| ...  |             |       |             |       |          |      | 
| ...  |             |       |             |       |          |      | 
| ...  |  NC         |  ...  |  target1    |  ...  |  ct25    |  ... | 
| ...  |  NC         |  ...  |  target1    |  ...  |  ct26    |  ... | 
| ...  |  NC         |  ...  |  target2    |  ...  |  ct27    |  ... | 
| ...  |  NC         |  ...  |  target2    |  ...  |  ct28    |  ... | 
| ...  |  NC         |  ...  |  target3    |  ...  |  ct29    |  ... | 
| ...  |  NC         |  ...  |  target3    |  ...  |  ct30    |  ... | 


Unfortunately, the data we want to have in hand is the structure like this:

|          |           |           |           |           |         |      | 
|----------|-----------|-----------|-----------|-----------|---------|------| 
| target1  |  target1  |  target2  |  target2  |  target3  | target3 |      | 
| PC       |  ct1      |  ct2      |  ct3      |  ct4      |  ct5    | ct6  | 
| sample1  |  ct7      |  ct8      |  ct9      |  ct10     |  ct11   | ct12 | 
| sample2  |  ct13     |  ct14     |  ct15     |  ct16     |  ct17   | ct18 | 
| sample3  |  ct19     |  ct20     |  c21      |  ct22     |  ct23   | ct24 | 
| ...      |           |           |           |           |         |      | 
| ...      |           |           |           |           |         |      | 
| NC       |  ct25     |  ct26     |  ct27     |  ct28     |  ct29   | ct30 | 

So, this is what we would do:
read data in each excel file, analysis it, represent it in a new excel sheet.
this will save a lot of Copy-and-Paste time, and your keyboard.

# Usage:
'''Python3 analysis_taqman_data.py <your_path_to_.xls_file(s)>'''

## Package needed:
xlrd, xlwt, sys, time

output will be a excel with file name of your local date-and-time.xls
