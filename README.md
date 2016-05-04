# survey_monkey
## 2016_05_04

```
usage: SM_to_CCNC.py [-h] [-i INPUTEXCEL [INPUTEXCEL ...]] [-g GROUP]
                     [-t TEMPLATE] [-st SURVEYTEMPLATE] [-o OUTPUT]

SM_to_CCNC.py : Saves the Survey monkey data to CCNC format
========================================

optional arguments:
  -h, --help            show this help message and exit
  -i INPUTEXCEL [INPUTEXCEL ...], --inputExcel INPUTEXCEL [INPUTEXCEL ...]
                        Survey Monkey exported data files
  -g GROUP, --group GROUP
                        Group fo the subject
  -t TEMPLATE, --template TEMPLATE
                        Excel template with all formulas
  -st SURVEYTEMPLATE, --surveyTemplate SURVEYTEMPLATE
                        Survey monkey template with question name column
  -o OUTPUT, --output OUTPUT
                        Excel ouput
```


## Example 
```
python /ccnc_bin/survery_monkey/SM_to_CCNC.py -i Sheet_1.xls Sheet_2.xls Sheet_3.xls
```
