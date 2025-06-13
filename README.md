# go-excel

One set of code to read xls and xlsx at the same time
Usage tutorial

Supports reading files and reading file streams
```
    Import(filename string)
    ImportByReader(reader io.Reader, fileType string)


    rows, err := file.GetRowIndex(0)
    if err != nil {
    return
    }
    
    for _, row := range rows {
    ...
    }
```
or
```
    for d, s := range file.GetSheetsName() {
        rows, err := file.GetRow(s)
        if err != nil {
            return
        }
        for _, row := range rows { 
            --your code
        }
    }
```
