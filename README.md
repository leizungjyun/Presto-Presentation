# Piano Group Toolbox
A toolbox for piano group of ZJU Graduate Art Troupe.


## Requried
```
pip install pandas openpyxl
brew install marp-cli
```



## Excel to Program Presentation
After collecting the information of the program, you can use the following command to generate the program presentation.

```bash
./xl2pdf <input-file> <output>
```

`input-file` is a `.xlsx` file, which contains the information of the program. `output` is the filename without extension. For example,

```bash
./xl2pdf example.xlsx program
```

Two files `program.md` and `program.pdf` will be generated. Use `program.pdf` to do presentation. You can also modify `program.md` to change the content of the presentation. Then, you can do

```bash
./md2pdf program.md 
```

