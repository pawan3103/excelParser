#README#

#Python 3.6.1

- Excel Parser with disk-file as well as with Bytes and unit testing for the parser.

#What is this repo#

- This is the code to parse excel data, which will have item and availableStock column.
It reads and parse it to namedtuple, right now we are catching only `item` and `availableStock`
if there are more parameters, then we can handle them as well, also we can trigger somer alerts
if they are missing, as per our requirements.
- In this demo we are dealing with disk-file, but I have shown you how you can deal with bytes as well by converting disk-file to bytes.
- Right now we are dealing everything with simple python script inclusing testing, later we will switch to some framework like django/flask as per requirement. 

#Sample excel file is also kept in BASE_DIR `sample.xlsx` for reference.

- `sample.xlsx` file is used for testing the code.
- Code is flak8 tested.

### How do I get set up? ###

- Create virtual environment and install requirements from `production.txt` for prod env
  and `test.txt` for testing.
- `base.txt` is common requirements for both prod and test env.

#Run the Code

- CD to BASE_DIR ie.. parser folder.
- Run `python excel_parser.py /path/to/excel/file`.
- To see the output in text file run `python excel_parser.py /path/to/excel/file > output.txt`
  Example: `python excel_parser.py 'sample.xlsx' > output.txt`
- You will see a file `output.txt` with output data.

#Run Test Case

- CD to BASE_DIR ie.. in parser folder.
- `python test.py` if there is any error it will throw it on console.
- or you can try `python test.py > test.txt` to see the output.
- To test flake8 `flake8 excel_parser.py` or simple `flake8`.