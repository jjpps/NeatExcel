# NeatExcel
Excel Reader Dll

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
