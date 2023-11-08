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
#### Python Packages
```bash
pip install pandas openpyxl
```

#### Marp CLI
See installation of [Marp CLI](https://github.com/marp-team/marp-cli) on your platform.

```bash
brew install marp-cli # on mac
```

## Usage

### Excel to Concert Presentation
After collecting the information of the program into an Excel file, you can use the following command to generate the concert presentation.

```bash
python presto -s [input-file]
# or
python presto --slides [input-file]
```
The `input-file` is optional, and the default value is `example.xlsx`.  `input-file` is a `.xlsx` file, which contains the information of the program.  The generated presentation will be saved in `slides.pdf`.  Another file `slides.md`, which is an intermediate result will also be generated.  You can also manully modify `slides.md` to change the content of the presentation. Then, you can do
```bash
python presto -m [slides.md]
```


### Excel to Concert Program

```bash
python presto -p  [input-file]
# or
python presto --program  [input-file]
```
The generated program will be saved in `program.pdf`.

### Excel to both Concert Presentation and Program

```bash
python presto -a  [input-file]
# or
python presto --all  [input-file]
# or
python presto -p -s [input-file]
```


