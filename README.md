# estimator

Estimator.py is a Python script designed to streamline raw material estimation and job number generation for a manufacturing processes.

## Features

- Efficiently estimate raw material requirements based on user input
- Creates instantaneous bill of material
- Generate unique job numbers for each estimation
- User-friendly interface for easy data entry and processing
- Error handling for invalid inputs and range checks
- Support for various job types and printing options
- Integration with Excel spreadsheets for data storage and analysis
- Easy customization and extension for specific manufacturing needs

## Usage
1 Install Python and the required dependencies.\
2 Run estimator.py using a Python interpreter.\
3 Enter the necessary job details and click the "Process" button.\
4 Review the estimated raw material requirements and job number generated.\
5 Save or export the data for further analysis or use in manufacturing processes.\

### OR

1 Export the file as an executable (.exe file) via Pyinstaller by using the '''pyinstaller --onefile -w --add-data="template.xlsx;estimator" estimator.py'''
  
## License
This project is licensed under the MIT License.

## Acknowledgements
[Openpyxl](https://pypi.org/project/openpyxl/) - Python library for interacting with Excel files
[Tkinter](https://docs.python.org/3/library/tkinter.html) - Python's standard GUI package
[PyInstaller](https://pyinstaller.org/en/stable/) - Python packaging and executable creation tool
