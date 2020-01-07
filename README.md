# NeatExcel


### Como Escrever

```
Excel excel = new Excel();
DataTable table = new DataTable();
table.Columns.Add("Dosage", typeof(int));
table.Columns.Add("Drug", typeof(string));
table.Columns.Add("Patient", typeof(string));
table.Columns.Add("Date", typeof(DateTime));

// Here we add five DataRows.
table.Rows.Add(25, "Indocin", "David", DateTime.Now);
table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);

DataSet ds = new DataSet();
ds.Tables.Add(table);
excel.WriteXLSX(ds, $"{path}\\Arquivo.xlsx");

```
### Como Ler
```

Excel excel = new Excel();
var excelLoaded = excel.LoadXLS(path);
foreach (DataRow dr in excelLoaded.Tables[1].Rows)
{
   dr["ID_TIPOAJUSTE"].ToString()
   //sw.WriteLine($"{item}");
}

```
