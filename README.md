# go-excel

一套代码同时读取xls和xlsx
使用教程

    rows, err := file.GetRowIndex(0)
    if err != nil {
    return
    }
    
    for _, row := range rows {
    ...
    }

or

    for d, s := range file.GetSheetsName() {
        rows, err := file.GetRow(s)
        if err != nil {
            return
        }
        for _, row := range rows { 
            --your code
        }
    }