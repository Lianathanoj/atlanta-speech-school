# Atlanta Speech School

This repository contains scripts for sanitizing Excel workbooks, in order to make the data work with Tableau.

### How to Use
The first thing to do is to clone this repository.

There are two ways to run this script:
* **Easy Use:** Double click the `excel-cleanup.exe` file. When an input window pops up, choose the excel spreadsheet that needs to be cleaned up.
* **Technical:** There are a couple of dependencies that need to be installed, which can be done by running `pip install -r requirements.txt`. Once all of the dependencies have been installed, run `python excel-cleanup.py`. This will open up the input window, into which you can input a file.

After running the script, a new file called `Data_MODIFIED.xlsx` will be created, containing geocoded, properly formatted data for use with Tableau. Now, simply import this file into Tableau, and you're done!

### For Developers
If you need to modify anything in the script, make sure to recreate the `.exe` file with [Pyinstaller](http://www.pyinstaller.org/).
1. Install Pyinstaller with `pip install pyinstaller`
2. Go into the project directory and type `pyinstaller --debug --onefile --noupx excel-cleanup.py`

Doing this will create a new `.exe` which can be run.
