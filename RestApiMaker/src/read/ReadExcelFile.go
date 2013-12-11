package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
)

func main() {
 	var xlsxFile *xlsx.File
	var error error
	xlsxFile, error = xlsx.OpenFile("d:\\test2.xlsx")
	if error != nil {
		fmt.Println(error.Error)
		fmt.Errorf(error.Error())
		return
	}
    if xlsxFile == nil {
		fmt.Errorf("OpenFile returned nil FileInterface without generating an os.Error")
		return
	}
    for _, sheet := range xlsxFile.Sheets {
        for _, row := range sheet.Rows {
//        	fmt.Printf("row name : %s\n", row.String())
			var i int
			i = 65		// A
            for _, cell := range row.Cells {
            	if(len(cell.String()) != 0) {
	                if(i<91) {
	                	fmt.Printf("sheet: %s, rownum: %s, column: %c, value: %s\n", sheet.String(), row.String(), i, cell.String())
	                	i += 1
	                } else {
	                	if(i==91) {
	                		i=65
	                	}
	                	fmt.Printf("sheet: %s, rownum: %s, column: A%c, value: %s\n", sheet.String(), row.String(), i, cell.String())
	                	// AA, BB, CC 형식으로 표현
	                }
                } else {
                	 if(i<91) {
	                	i += 1
	                } else {
	                	
	                	// AA, BB, CC 형식으로 표현
	                }
                }
            }
        }
    }

}