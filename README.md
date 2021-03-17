# iziexcel

iziexcel is a simplified package for created basic excel files using [excelize](https://github.com/360EntSecGroup-Skylar/excelize)

#### Use [excelize](https://github.com/360EntSecGroup-Skylar/excelize) for adding chart, picture and detailed usage.

```
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

err = iziexcel.Save(f, "test.xlsx") // or .xls
if err != nil {
	panic(err)
}
```