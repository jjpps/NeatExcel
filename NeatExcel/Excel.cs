using Microsoft.VisualBasic.FileIO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NeatExcel
{
    public class Excel
    {
        private HSSFWorkbook hssfworkbook;
        private XSSFWorkbook xssfworkbook;

        private readonly Hashtable hashCellStyles = new Hashtable();
        private readonly Hashtable hashFontStyles = new Hashtable();

        public DataTable LoadExcel(string path, string table)
        {
            return LoadExcel(path, table, 0);
        }

        public DataTable LoadExcel(string path, string table, int countColumns)
        {
            FileInfo f = new FileInfo(path);

            if (f.Extension == ".xls")
            {
                return LoadXLS(path, table, countColumns);
            }
            else
            {
                return LoadXLSX(path, table, countColumns);
            }
        }

        private DataTable LoadXLSX(string path, string table, int countColumns)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xssfworkbook = new XSSFWorkbook(file);
            }

            ISheet sheet = null;

            bool found = false;

            int c = xssfworkbook.Count();
            for (int i = 0; i < c; i++)
            {
                sheet = xssfworkbook.GetSheetAt(i);

                if (sheet.SheetName.ToLower(CultureInfo.InvariantCulture) == table.ToLower(CultureInfo.InvariantCulture))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                return new DataTable();
            }

            return ReadSheet(sheet, countColumns);
        }

        private DataTable ReadSheet(ISheet sheet, int countColumns)
        {
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            DataTable dt = new DataTable(sheet.SheetName);

            rows.MoveNext();
            IRow rowHeader = (IRow)rows.Current;

            if (countColumns == 0)
            {
                countColumns = rowHeader.LastCellNum;
            }

            for (int i = 0; i < countColumns; i++)
            {
                ICell cell = rowHeader.GetCell(i);

                if (cell != null)
                {
                    dt.Columns.Add(cell.ToString());
                }
                else
                {
                    dt.Columns.Add("");
                }
            }

            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);


                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                            case CellType.Formula:
                                {
                                    try
                                    {
                                        dr[i] = cell.NumericCellValue;
                                    }
                                    catch
                                    {
                                        dr[i] = cell.StringCellValue;
                                    }
                                    break;
                                }
                            default:
                                {
                                    dr[i] = cell.ToString();
                                    break;
                                }
                        }
                    }
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private DataTable LoadXLS(string path, string table, int countColumns)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }

            ISheet sheet = null;

            bool found = false;

            for (int i = 0; i < 10; i++)
            {
                sheet = hssfworkbook.GetSheetAt(i);

                if (sheet.SheetName.ToLower(CultureInfo.InvariantCulture) == table.ToLower(CultureInfo.InvariantCulture))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                return new DataTable();
            }

            return ReadSheet(sheet, countColumns);
        }

        public DataSet LoadXLS(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }

            DataSet ds = new DataSet();

            for (int x = 0; x < hssfworkbook.NumberOfSheets; x++)
            {
                ISheet sheet = hssfworkbook.GetSheetAt(x);

                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

                DataTable dt = new DataTable(sheet.SheetName);

                rows.MoveNext();
                IRow rowHeader = (HSSFRow)rows.Current;

                for (int i = 0; i < rowHeader.LastCellNum; i++)
                {
                    ICell cell = rowHeader.GetCell(i);

                    if (cell != null)
                    {
                        dt.Columns.Add(cell.ToString());
                    }
                    else
                    {
                        dt.Columns.Add("");
                    }
                }

                while (rows.MoveNext())
                {
                    IRow row = (HSSFRow)rows.Current;
                    DataRow dr = dt.NewRow();

                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        ICell cell = row.GetCell(i);


                        if (cell == null)
                        {
                            dr[i] = null;
                        }
                        else
                        {
                            dr[i] = cell.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                }

                ds.Tables.Add(dt);
            }

            return ds;
        }

        public DataSet LoadXLSX(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                xssfworkbook = new XSSFWorkbook(file);
            }

            DataSet ds = new DataSet();

            for (int x = 0; x < xssfworkbook.NumberOfSheets; x++)
            {
                ISheet sheet = xssfworkbook.GetSheetAt(x);

                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

                DataTable dt = new DataTable(sheet.SheetName);

                rows.MoveNext();
                IRow rowHeader = (IRow)rows.Current;

                for (int i = 0; i < rowHeader.LastCellNum; i++)
                {
                    ICell cell = rowHeader.GetCell(i);

                    if (cell != null)
                    {
                        dt.Columns.Add(cell.ToString().TrimEnd());
                    }
                    else
                    {
                        dt.Columns.Add("");
                    }
                }

                while (rows.MoveNext())
                {
                    IRow row = (IRow)rows.Current;
                    DataRow dr = dt.NewRow();

                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        ICell cell = row.GetCell(i);


                        if (cell == null)
                        {
                            dr[i] = null;
                        }
                        else
                        {
                            dr[i] = cell.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                }

                ds.Tables.Add(dt);
            }

            return ds;
        }

        public DataSet LoadCSV(string path)
        {
            TextFieldParser textReader = new TextFieldParser(path, Encoding.GetEncoding(1252));

            textReader.SetDelimiters(new string[] { ";", "\t" });

            string[] fields = textReader.ReadFields();

            DataSet ds = new DataSet();

            DataTable dt = new DataTable("Table1");

            foreach (string field in fields)
            {
                dt.Columns.Add(field);
            }

            while (!textReader.EndOfData)
            {
                string[] dados = textReader.ReadFields();

                DataRow dr = dt.NewRow();

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    dr[i] = dados[i];
                }

                dt.Rows.Add(dr);
            }

            ds.Tables.Add(dt);

            return ds;
        }

        public void WriteXLS(DataSet ds, string path)
        {
            WriteXLS(ds, path, Estilo.Padrao);
        }

        public void WriteXLSX(DataSet ds, string path)
        {
            WriteXLSX(ds, path, Estilo.Padrao);
        }

        public void WriteXLS(DataSet ds, string path, Estilo estilo)
        {
            //Cria o arquivo Excel
            IWorkbook workbook = new HSSFWorkbook();

            foreach (DataTable dt in ds.Tables)
            {
                //Cria uma planilha dentro do arquivo e define o nome
                ISheet sheet = workbook.CreateSheet(dt.TableName);

                int colIndex = 0;

                IRow linha = sheet.CreateRow(0);

                foreach (DataColumn dc in dt.Columns)
                {
                    ICell cell = linha.CreateCell(colIndex);
                    cell.CellStyle = BuscaEstiloCabecalho(estilo, workbook);
                    cell.SetCellValue(dc.ColumnName);

                    colIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    linha = sheet.CreateRow(rowIndex);

                    colIndex = 0;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        ICell cell = linha.CreateCell(colIndex);
                        cell.CellStyle = BuscaEstiloLinhas(estilo, workbook);
                        cell.SetCellValue(dr[dc].ToString());

                        colIndex++;
                    }

                    rowIndex++;
                }

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }

            //salva o arquivo
            FileStream file;

            file = new FileStream(path, FileMode.Create);

            workbook.Write(file);
            file.Close();
        }

        public void WriteXLSX(DataSet ds, string path, Estilo estilo)
        {
            //Cria o arquivo Excel
            IWorkbook workbook = new XSSFWorkbook();

            foreach (DataTable dt in ds.Tables)
            {
                //Cria uma planilha dentro do arquivo e define o nome
                ISheet sheet = workbook.CreateSheet(dt.TableName);

                int colIndex = 0;

                IRow linha = sheet.CreateRow(0);

                foreach (DataColumn dc in dt.Columns)
                {
                    ICell cell = linha.CreateCell(colIndex);
                    cell.CellStyle = BuscaEstiloCabecalho(estilo, workbook);
                    cell.SetCellValue(dc.ColumnName);

                    colIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    linha = sheet.CreateRow(rowIndex);

                    colIndex = 0;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        ICell cell = linha.CreateCell(colIndex);
                        cell.CellStyle = BuscaEstiloLinhas(estilo, workbook);
                        cell.SetCellValue(dr[dc].ToString());

                        colIndex++;
                    }

                    rowIndex++;
                }

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }


            CultureInfo ci = new CultureInfo("en-US");
            Thread.CurrentThread.CurrentCulture = ci;
            Thread.CurrentThread.CurrentUICulture = ci;

            //salva o arquivo
            FileStream file;

            file = new FileStream(path, FileMode.Create, FileAccess.Write);

            workbook.Write(file);
            file.Close();

        }

        private ICellStyle BuscaEstiloCabecalho(Estilo estilo, IWorkbook workbook)
        {
            switch (estilo)
            {
                case Estilo.Padrao:
                    {
                        //Cria o estilo pra usar no cabeçalho
                        ICellStyle stlCabecalho = workbook.CreateCellStyle();
                        stlCabecalho.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                        stlCabecalho.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                        stlCabecalho.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                        stlCabecalho.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                        stlCabecalho.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                        stlCabecalho.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                        stlCabecalho.SetFont(BuscaFonteCabecalho(estilo, workbook));

                        return stlCabecalho;
                    }
                case Estilo.SemBorda:
                    {
                        //Cria o estilo pra usar no cabeçalho
                        ICellStyle stlCabecalho = workbook.CreateCellStyle();
                        stlCabecalho.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
                        stlCabecalho.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
                        stlCabecalho.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
                        stlCabecalho.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
                        stlCabecalho.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                        stlCabecalho.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                        stlCabecalho.SetFont(BuscaFonteCabecalho(estilo, workbook));

                        return stlCabecalho;
                    }
                default: return null;
            }
        }

        private ICellStyle BuscaEstiloLinhas(Estilo estilo, IWorkbook workbook)
        {
            switch (estilo)
            {
                case Estilo.Padrao:
                    {
                        ICellStyle stlLinhas = (ICellStyle)this.hashCellStyles[Estilo.Padrao];

                        if (stlLinhas == null)
                        {
                            stlLinhas = workbook.CreateCellStyle();
                            stlLinhas.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            stlLinhas.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                            stlLinhas.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                            stlLinhas.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            stlLinhas.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                            stlLinhas.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                            stlLinhas.SetFont(BuscaFonteLinhas(estilo, workbook));

                            this.hashCellStyles.Add(Estilo.Padrao, stlLinhas);
                        }
                        return stlLinhas;
                    }
                case Estilo.SemBorda:
                    {
                        ICellStyle stlLinhas = (ICellStyle)this.hashCellStyles[Estilo.Padrao];

                        if (stlLinhas == null)
                        {
                            stlLinhas = workbook.CreateCellStyle();
                            stlLinhas.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
                            stlLinhas.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
                            stlLinhas.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
                            stlLinhas.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
                            stlLinhas.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                            stlLinhas.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                            stlLinhas.SetFont(BuscaFonteLinhas(estilo, workbook));

                            this.hashCellStyles.Add(Estilo.Padrao, stlLinhas);
                        }
                        return stlLinhas;
                    }
                default: return null;
            }
        }

        private IFont BuscaFonteCabecalho(Estilo estilo, IWorkbook workbook)
        {
            switch (estilo)
            {
                case Estilo.Padrao:
                case Estilo.SemBorda:
                    {
                        //Cria fonte pra ser usada no cabeçalho
                        IFont fonteCabecalho = workbook.CreateFont();
                        fonteCabecalho.FontHeightInPoints = 10;
                        fonteCabecalho.FontName = "Arial";
                        fonteCabecalho.Boldweight = (short)FontBoldWeight.Bold;
                        fonteCabecalho.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;

                        return fonteCabecalho;
                    }
                default: return null;
            }
        }

        private IFont BuscaFonteLinhas(Estilo estilo, IWorkbook workbook)
        {
            switch (estilo)
            {
                case Estilo.Padrao:
                case Estilo.SemBorda:
                    {

                        IFont fonteLista = (IFont)this.hashFontStyles[Estilo.Padrao];

                        if (fonteLista == null)
                        {
                            //Cria fonte pra ser usada na lista
                            fonteLista = workbook.CreateFont();
                            fonteLista.FontHeightInPoints = 10;
                            fonteLista.FontName = "Arial";
                            fonteLista.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;

                            this.hashFontStyles.Add(Estilo.Padrao, fonteLista);
                        }

                        return fonteLista;
                    }
                default: return null;
            }
        }
    }
}
