package iziexcel

import (
	"errors"
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type NewExcel struct {
	Headers   []string
	Rows      [][]interface{}
	SheetName string
}

func (prepareExcel *NewExcel) Create() (*excelize.File, error) {
	if prepareExcel.SheetName == "" {
		prepareExcel.SheetName = "Sheet1"
	}

	f := excelize.NewFile()
	index := f.NewSheet(prepareExcel.SheetName)
	f.SetActiveSheet(index)

	if len(prepareExcel.Headers) > 0 {
		for index, header := range prepareExcel.Headers {
			column, err := excelize.ColumnNumberToName(index + 1)
			if err != nil {
				return nil, err
			}
			err = f.SetCellValue("Sheet1", fmt.Sprintf("%s1", column), header)
			if err != nil {
				return nil, err
			}
		}
	}
	// if headers len is 0 file started from 1. row, if headers len > 0 file started from 2. row
	setGap := 1
	if len(prepareExcel.Headers) > 0 {
		setGap = 2
	}
	for rowIndex, row := range prepareExcel.Rows {
		for columnIndex, value := range row {
			columnIndex++
			columnName, err := excelize.ColumnNumberToName(columnIndex)
			if err != nil {
				return nil, err
			}
			err = f.SetCellValue(prepareExcel.SheetName, fmt.Sprintf("%s%d", columnName, rowIndex+setGap), value)
			if err != nil {
				return nil, err
			}
		}
	}

	return f, nil
}

func Save(file *excelize.File, fileName string) error {
	if file == nil {
		return errors.New("file can not be nil")
	}
	return file.SaveAs(fileName)
}
