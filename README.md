# xlwings-watch-license

Install python 3: https://www.python.org/downloads/

Install packages:
```sh
pip install xlwings
pip install requests
```

Change some values:
```python
# If config page is http://localhost:38080/page/index.html, server address is:
DLS_SERVER = 'http://localhost:38080/' 
# license item ids
LICENSE_ITEMS = ['123456','234567']
# can be found in F12 Console -> Network -> Fetch/XHR ->
# http://localhost:38080/license/item/ -> Request Headers
DynamsoftLTSTokenV2 = '?????' 
```

Run:
```sh
python ./update-excel.py
```

The first execution will generate `watch-license.xlsx` in the folder of `update-excel.py`.

The next time, you should execute this command when the `watch-license.xlsx` is closed.

Each execution will add a column record.

Please modify the script according to your needs.
