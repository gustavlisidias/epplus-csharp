// Decompiled with JetBrains decompiler
// Type: EpPlusGx16.AasExcelGx16
// Assembly: EpPlusGx16, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 2A6B3DBC-342E-487B-8574-2C4C616ADA8A
// Assembly location: C:\Users\gustavo.dias\Downloads\EpPlusGx16.dll

using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;

namespace EpPlusGx16
{
    public class AasExcelGx16
    {
        private ExcelPackage excelPackage;
        private ExcelWorksheet worksheet;

        public void Open(string Directory, string Name, string Extension)
        {
            string fullPath = Path.Combine(Directory, Name + "." + Extension);

            if (File.Exists(fullPath))
            {
                // Abrir a instancia com o arquivo, caso encontrado
                this.excelPackage = new ExcelPackage(new FileInfo(fullPath));
            }
            else
            {
                // Chamar função para criar o arquivo
                Create(fullPath);
            }
        }

        public void OpenFromTemplate(string templatePath, string Directory, string Name, string Extension)
        {
            string fullPath = Path.Combine(Directory, Name + "." + Extension);

            if (File.Exists(fullPath))
            {
                // Abrir a instancia com o arquivo, caso encontrado
                this.excelPackage = new ExcelPackage(new FileInfo(fullPath));
            }
            else
            {
                // Chamar função para criar o arquivo a partir de um template, caso não encontrado
                CreateFromTemplate(templatePath, fullPath);
            }
        }

        private void Create(string fullPath)
        {
            // Criar nova instancia
            this.excelPackage = new ExcelPackage();

            // Adiciona uma nova sheet
            this.worksheet = this.excelPackage.Workbook.Worksheets.Add("Planilha1");

            // Salvar a instância do ExcelPackage no novo arquivo
            this.excelPackage.SaveAs(new FileInfo(fullPath));

            // Abrir a instancia com o arquivo que foi salvo
            this.excelPackage = new ExcelPackage(new FileInfo(fullPath));
        }

        private void CreateFromTemplate(string templatePath, string fullPath)
        {
            if (File.Exists(templatePath))
            {
                // Carregar o arquivo Excel template
                using (var templatePackage = new ExcelPackage(new FileInfo(templatePath)))
                {
                    // Clonar a planilha do template para o novo arquivo
                    this.excelPackage = new ExcelPackage(templatePackage.Stream);
                }

                // Salvar a instância do ExcelPackage no novo arquivo
                this.excelPackage.SaveAs(new FileInfo(fullPath));

                // Abrir a instancia com o arquivo que foi salvo
                this.excelPackage = new ExcelPackage(new FileInfo(fullPath));
            }
            else
            {
                // Caso não seja encontrado o template, throw Error
                throw new FileNotFoundException("O arquivo de template não foi encontrado.", templatePath);
            }
        }

        //private ExcelPackage excelPackage = new ExcelPackage();
        //private ExcelWorksheet worksheet;

        //public void Open(string Directory, string Name, string Extension)
        //{
        //    this.excelPackage = new ExcelPackage(new FileInfo(Directory + "/" + Name + "." + Extension));
        //}

        //public void OpenFromTemplate(string templatePath, string outputPath)
        //{
        //    Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(templatePath);
        //    this.excelPackage = new ExcelPackage(stream);

        //    FileInfo outputFile = new FileInfo(outputPath);
        //    this.excelPackage.SaveAs(outputFile);

        //    excelPackage.Dispose();
        //}

        //public void OpenFromTemplate1(string templatePath, string outputPath)
        //{
        //    // Carrega o arquivo Excel template
        //    ExcelPackage templatePackage = new ExcelPackage(new FileInfo(templatePath));

        //    // Criação de um novo arquivo
        //    string newFilePath = Path.Combine(Path.GetDirectoryName(templatePath), outputPath);
        //    FileInfo newFile = new FileInfo(newFilePath);

        //    // Cópia dos dados
        //    File.Copy(templatePath, newFilePath, true);

        //    // Atribuo ao this.excelPackage
        //    this.excelPackage = new ExcelPackage(newFile);
        //}

        //public void GenerateFromTemplate()
        //{
        //    string saida = this.excelPackage.File.FullName;

        //    Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(Template);
        //    this.excelPackage = new ExcelPackage(stream);

        //    FileInfo outputFile = new FileInfo(saida);
        //    this.excelPackage.SaveAs(outputFile);

        //    excelPackage.Dispose();
        //}

        public void SelectSheet(string NameSheet)
        {
            var TabSheet = this.excelPackage.Workbook.Worksheets[NameSheet];

            if (TabSheet != null)
            {
                // Selecionar a sheet, caso exista
                this.worksheet = this.excelPackage.Workbook.Worksheets[NameSheet];
            }
            else
            {
                // Chamar função para criar nova sheet, caso não exista
                CreateSheet(NameSheet);
            }
        }

        public void CreateSheet(string NameSheet) => this.worksheet = this.excelPackage.Workbook.Worksheets.Add(NameSheet);

        public void CellsHorizontalSize(int ColumnStart, int Size) => this.worksheet.Column(ColumnStart).Width = (double)Size;

        public void CellsVerticalSize(int RowStart, int Size) => this.worksheet.Row(RowStart).Height = (double)Size;

        public void Cells(int RowStart, int ColumnStart, string Text) => this.worksheet.Cells[RowStart, ColumnStart].Value = (object)Text;

        public void Cells1(int RowStart, int ColumnStart, int RowLast, int ColumnLast, string Text) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Value = (object)Text;

        public void CellsNumber(int RowStart, int ColumnStart, double Numero) => this.worksheet.Cells[RowStart, ColumnStart].Value = (object)Numero;

        public void CellsNumber1(int RowStart, int ColumnStart, int RowLast, int ColumnLast, double Numero) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Value = (object)Numero;

        public void CellsHorizontalCenter(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        public void CellsHorizontalCenter1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        public void CellsHorizontalLeft(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

        public void CellsHorizontalLeft1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

        public void CellsHorizontalRight(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        public void CellsHorizontalRight1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        public void CellsHorizontalJustify(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

        public void CellsHorizontalJustify1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

        public void CellsVerticalTop(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

        public void CellsVerticalTop1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

        public void CellsVerticalCenter(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        public void CellsVerticalCenter1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        public void CellsVerticalBottom(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

        public void CellsVerticalBottom1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

        public void CellsVerticalJustify(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.VerticalAlignment = ExcelVerticalAlignment.Justify;

        public void CellsVerticalJustify1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.VerticalAlignment = ExcelVerticalAlignment.Justify;

        public void CellsFormula(int RowStart, int ColumnStart, string Formula) => this.worksheet.Cells[RowStart, ColumnStart].Formula = Formula;

        public void CellsFormula1(int RowStart, int ColumnStart, string Formula) => this.worksheet.Cells[RowStart, ColumnStart].FormulaR1C1 = Formula;

        public void CellsBold(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.Font.Bold = true;

        public void CellsBold1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Bold = true;

        public void CellsItalic(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.Font.Italic = true;

        public void CellsItalic1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Italic = true;

        public void CellsUnderLine(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.Font.UnderLine = true;

        public void CellsUnderLine1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.UnderLine = true;

        public void CellsFontColor(int RowStart, int ColumnStart, string Color) => this.worksheet.Cells[RowStart, ColumnStart].Style.Font.Color.SetColor(ColorTranslator.FromHtml(Color));

        public void CellsFontColor1(int RowStart, int ColumnStart, int RowLast, int ColumnLast, string Color) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Color.SetColor(ColorTranslator.FromHtml(Color));

        public void CellsRotation(int RowStart, int ColumnStart, int Rotacion) => this.worksheet.Cells[RowStart, ColumnStart].Style.TextRotation = Rotacion;

        public void CellsFontSize(int RowStart, int ColumnStart, int Size) => this.worksheet.Cells[RowStart, ColumnStart].Style.Font.Size = (float)Size;

        public void CellsFontSize1(int RowStart, int ColumnStart, int RowLast, int ColumnLast, int Size) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Size = (float)Size;

        public void CellsFontName(int RowStart, int ColumnStart, string Font) => this.worksheet.Cells[RowStart, ColumnStart].Style.Font.Name = Font;

        public void CellsFontName(int RowStart, int ColumnStart, int RowLast, int ColumnLast, string Font) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Font.Name = Font;

        public void CellsCombine(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Merge = true;

        public void CellsWrapText(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.WrapText = true;

        public void CellsFormatNum(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "0.00";

        public void CellsFormatNum1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = "0.00";

        public void CellsFormatNumCom(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "#,##0.00";

        public void CellsFormatNumCom1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = "#,##0.00";

        public void CellsFormatNumB(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "#";

        public void CellsFormatNumB1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = "#";

        public void CellsFormatDate(int RowStart, int ColumnStart) => this.worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = "mm-dd-yyyy";

        public void CellsFormatDate1(int RowStart, int ColumnStart, int RowLast, int ColumnLast) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = "mm-dd-yyyy";

        public void CellsFormatPersonalizado(int RowStart, int ColumnStart, string formato) => this.worksheet.Cells[RowStart, ColumnStart].Style.Numberformat.Format = formato;

        public void CellsFormatPersonalizado1(int RowStart, int ColumnStart, int RowLast, int ColumnLast, string formato) => this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Numberformat.Format = formato;

        public void CellsFormulaOper(int RowStart, int ColumnStart, int RowFirstCell, int ColumnFirstCell, int RowSecundCell, int ColumnSecundCell, string Operation)
        {
            this.worksheet.Cells[RowStart, ColumnStart].Formula = "(" + this.worksheet.Cells[RowFirstCell, ColumnFirstCell].Address + Operation + this.worksheet.Cells[RowSecundCell, ColumnSecundCell].Address + ")";
        }

        public void CellsBord(int RowStart, int ColumnStart)
        {
            this.worksheet.Cells[RowStart, ColumnStart].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            this.worksheet.Cells[RowStart, ColumnStart].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            this.worksheet.Cells[RowStart, ColumnStart].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            this.worksheet.Cells[RowStart, ColumnStart].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        public void CellsBord1(int RowStart, int ColumnStart, int RowLast, int ColumnLast)
        {
            this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        public void CellsColor(int RowStart, int ColumnStart, string Color)
        {
            this.worksheet.Cells[RowStart, ColumnStart].Style.Fill.PatternType = ExcelFillStyle.Solid;
            this.worksheet.Cells[RowStart, ColumnStart].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(Color));
        }

        public void CellsColor1(int RowStart, int ColumnStart, int RowLast, int ColumnLast, string Color)
        {
            this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Fill.PatternType = ExcelFillStyle.Solid;
            this.worksheet.Cells[RowStart, ColumnStart, RowLast, ColumnLast].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(Color));
        }

        // Funcionalidades adicionais
        public void CellAddComment(int RowStart, int ColumnStart, string Text, string Author)
        {
            var cell = this.worksheet.Cells[RowStart, ColumnStart];
            var comment = cell.AddComment(Text + "\nDe: " + Author, Author);
            comment.AutoFit = true;
        }

        public void CellAddComment1(int RowStart, int ColumnStart, string Text, string Author, string NameSheet)
        {
            var TabSheet = this.excelPackage.Workbook.Worksheets[NameSheet];

            if (TabSheet != null)
            {
                this.worksheet = this.excelPackage.Workbook.Worksheets[NameSheet];
                var cell = this.worksheet.Cells[RowStart, ColumnStart];
                var comment = cell.AddComment(Text + "\nDe: " + Author, Author);
                comment.AutoFit = true;
            }
            else
            {
                this.worksheet = this.excelPackage.Workbook.Worksheets.Add(NameSheet);
                var cell = this.worksheet.Cells[RowStart, ColumnStart];
                var comment = cell.AddComment(Text + "\nDe: " + Author, Author);
                comment.AutoFit = true;
            }
        }

        public int ProcurarValorExato(string valorProcurado, string nameSheet, int indiceColuna)
        {
            var tabSheet = this.excelPackage.Workbook.Worksheets[nameSheet];

            if (tabSheet != null)
            {
                // Última linha da coluna A
                int ultimaLinha = tabSheet.Cells[tabSheet.Dimension.End.Row, indiceColuna].End.Row;

                for (int row = 1; row <= ultimaLinha; row++)
                {
                    if (tabSheet.Cells[row, indiceColuna].Text == valorProcurado)
                    {
                        // Retorna o número da linha onde o valor foi encontrado
                        return row;
                    }
                }

                // Retorna -1 caso o valor não seja encontrado na planilha
                return -1;
            }
            else
            {
                // Retorna 0 caso a planilha não seja encontrada
                return 0;
            }
        }

        public List<int> ProcurarValor(string valorProcurado, string nameSheet, int indiceColuna, bool addExato)
        {
            var tabSheet = this.excelPackage.Workbook.Worksheets[nameSheet];
            var resultados = new List<int>();

            if (tabSheet != null)
            {
                // Última linha da coluna A
                int ultimaLinha = tabSheet.Cells[tabSheet.Dimension.End.Row, indiceColuna].End.Row;

                for (int row = 1; row <= ultimaLinha; row++)
                {
                    if (addExato)
                    {
                        // Se contem ou se é igual é adicionado ao array
                        if (tabSheet.Cells[row, indiceColuna].Text == valorProcurado || tabSheet.Cells[row, indiceColuna].Text.Contains(valorProcurado))
                        {
                            resultados.Add(row);
                        }
                    }
                    else
                    {
                        // Apenas se contem é adicionado ao array
                        if (tabSheet.Cells[row, indiceColuna].Text.Contains(valorProcurado) && tabSheet.Cells[row, indiceColuna].Text != valorProcurado)
                        {
                            resultados.Add(row);
                        }
                    }
                }
            }

            // Retorna array com as linhas encontradas
            return resultados;
        }
        // Funcionalidades adicionais

        public void Grafica(int RowStartDesign, int ColumnStartDesign, int RowStartRead, int ColumnStartRead, int RowLastRead, int ColumnLastRead, int SizeWidth, int Height, string Name, string Title)
        {
            ExcelChart excelChart = this.worksheet.Drawings.AddChart(Name, eChartType.ColumnClustered);
            excelChart.Title.Text = Title;
            excelChart.SetPosition(RowStartDesign, 0, ColumnStartDesign, 0);
            excelChart.SetSize(SizeWidth, Height);
            excelChart.Legend.Remove();
            excelChart.Series.Add((ExcelRangeBase)this.worksheet.Cells[ColumnStartRead.ToString() + ":" + (object)ColumnLastRead], (ExcelRangeBase)this.worksheet.Cells[RowStartRead.ToString() + ":" + (object)RowLastRead]);
        }

        public void Image(int RowStart, int ColumnStart, string imagePath, int WidthSize, int HeightSize)
        {
            int num = new Random().Next(0, 100);
            Bitmap bitmap = new Bitmap(imagePath);
            if (bitmap == null)
                return;
            ExcelPicture excelPicture = this.worksheet.Drawings.AddPicture("Imagen" + (object)num, (System.Drawing.Image)bitmap);
            excelPicture.From.Column = ColumnStart;
            excelPicture.From.Row = RowStart;
            excelPicture.SetSize(WidthSize, HeightSize);
            excelPicture.From.ColumnOff = this.Pixel2MTU(2);
            excelPicture.From.RowOff = this.Pixel2MTU(2);
        }

        public int Pixel2MTU(int pixels) => pixels * 9525;

        public void Save() => this.excelPackage.Save();

        // Funcionalidades adicionais
        public void SaveClose()
        {
            this.excelPackage.Save();
            this.excelPackage.Dispose();
        }
    }
}
