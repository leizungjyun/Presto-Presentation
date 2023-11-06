# Piano Group Toolbox
A toolbox for piano group of ZJU Graduate Art Troupe.


## Requried
```
pip install pandas openpyxl
brew install marp-cli
```



## Excel to Concert Presentation
After collecting the information of the program, you can use the following command to generate the concert presentation.

```bash
./generate-presentation <input-file> 
```

`input-file` is a `.xlsx` file, which contains the information of the program. For example,

```bash
./generate-presentaion example.xlsx 
```

Two files `slides.md` and `slides.pdf` will be generated. Use `slides.pdf` to do presentation. You can also modify `slides.md` to change the content of the presentation. Then, you can do

```bash
./md2pdf slides.md 
```


## Excel to Concert Program

```bash
./generate-program <input-file> 
```
The generated program will be saved in `program.pdf`.