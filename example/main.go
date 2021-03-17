package main

import (
	"github.com/bdemirpolat/iziexcel"
)

func main() {
	headers := []string{"Cell A1", "Cell B1", "Cell C1"}
	var rows [][]interface{}
	rows = append(rows, []interface{}{"Cell A2", "Cell B2", "Cell C2"})
	rows = append(rows, []interface{}{"Cell A3", "Cell B3", "Cell C3"})
	excel := iziexcel.NewExcel{
		Headers: headers, // can be nil if you don't want to add headers
		Rows:    rows,
	}
	f, err := excel.Create()
	if err != nil {
		panic(err)
	}
	err = iziexcel.Save(f, "test.xlsx")
	if err != nil {
		panic(err)
	}

}
