# Presto Presentation
An Excel-to-PDF toolbox for the Piano Group of ZJU Graduate Art Troupe.

## Features
* Generate concert presentation slides from Excel file.
* Generate concert program pdf from Excel file.


## Requirements
* Python packages: pandas, openpyxl
* CLI interface of marp: marp-cli
* latex engine: xelatex

The presentation is based on Marp. The program is based on latex.

### Installation
```
pip install pandas openpyxl
brew install marp-cli
```


## Usage

### Excel to Concert Presentation
After collecting the information of the program into an Excel file, you can use the following command to generate the concert presentation.

```bash
./presto -s [input-file]
# or
./presto --slides [input-file]
```
The `input-file` is optional, and the default value is `example.xlsx`.  `input-file` is a `.xlsx` file, which contains the information of the program.  The generated presentation will be saved in `slides.pdf`.  Another file `slides.md`, which is an intermediate result will also be generated.  You can also manully modify `slides.md` to change the content of the presentation. Then, you can do
```bash
./md2pdf slides.md 
```


### Excel to Concert Program

```bash
./presto -p  [input-file]
# or
./presto --program  [input-file]
```
The generated program will be saved in `program.pdf`.

### Excel to both Concert Presentation and Program

```bash
./presto -a  [input-file]
# or
./presto --all  [input-file]
# or
./presto -p -s [input-file]
```


