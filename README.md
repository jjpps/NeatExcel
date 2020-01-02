# NeatExcel
Excel Reader Dll
###Como Escrever
```
DataTable table1 = new DataTable(GenereciType.Name);
foreach (var item in GenereciType.GetProperties().ToList().OrderBy(x => x.Name))
{
    table1.Columns.Add(item.Name);
}
if (typeof(T) == typeof(List<Clientes>))
{
    var lista = (List<Clientes>)obj;
    foreach (var item in lista)
    {
        table1.Rows.Add(new object[]
        {

        });
    }
}
```
### Como Ler
```
 foreach (DataRow dr in excelLoaded.Tables[i].Rows)
                    {
                        dr["ID_TIPOAJUSTE"].ToString()
                        //sw.WriteLine($"{item}");
                    }

```
