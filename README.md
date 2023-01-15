# Sheet converter


## Requirements

Those packages are essential to install in order for script to work

```    
python

# packages
pip install pandas
pip install openpyxl
```

## How to run

Minimal:

```
python ./sheet_converter.py --input-filename INPUT_FILENAME 
```

Full:

```
python sheet_converter.py 
    --input-filename INPUT_FILENAME 
    [--input-sheet INPUT_SHEET] 
    [--number-of-header NUMBER_OF_HEADER] 
    [--output-file OUTPUT_FILE]
```


## Troubleshooting

1. Try using `pip3` instead of `pip`
2. Tested on `python` version `3.9`
3. Maybe use `virtualenv`

