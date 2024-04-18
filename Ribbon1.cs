using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Authentication.ExtendedProtection.Configuration;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Diagnostics.Contracts;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ASRTookit_M
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        /// <summary>
        /// 元位数字格式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            selection.NumberFormat = "#,##0_);[Red](#,##0);-";
        }
        /// <summary>
        /// 反向
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            //System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("^(-?\\d+)(\\.\\d+)?$");
            //bool IsNumber = regex.IsMatch(selection.Value);
            int emptyCellCount = 0;
            foreach (Excel.Range row in selection.Rows)
            {
                foreach (Excel.Range cell in row.Cells)
                {
                    if (emptyCellCount >= 10000)
                    {
                        MessageBox.Show("Too Much Empty Cells", "10000+ emtpy cell found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (cell.Value == null)
                    {
                        emptyCellCount++;
                        continue;
                    }
                    if (cell.Value is double)
                    {
                        Debug.WriteLine("test");
                        if (cell.HasFormula)
                        {
                            String pureFormula = cell.Formula;
                            String pf = pureFormula.Substring(1);
                            if (pf.StartsWith("-(") && pf.EndsWith(")"))
                            {
                                String unNageativepf = pf.Substring(1);
                                cell.Formula = "=" + unNageativepf;
                                emptyCellCount = 0;
                                continue;
                            }
                            if (pf.StartsWith("(") && pf.EndsWith(")"))
                            {
                                cell.Formula = "=-" + pf;
                                emptyCellCount = 0;
                                continue;
                            }
                            cell.Formula = "=-(" + pf + ")";
                            emptyCellCount = 0;
                            continue;
                        }
                        cell.Formula = "=-(" + cell.Value + ")";
                        emptyCellCount = 0;
                        continue;
                    }
                    emptyCellCount = 0;
                    continue;
                }
            }
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            selection.NumberFormat = "#,##0.00_);[Red](#,##0.00);-";
        }

        private void Font_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            selection.Font.Name = "Microsoft YaHei";
            selection.Font.Size = 10;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            int chunkSize = 100; // 每次处理的行数或列数

            // 按行分段处理
            for (int startRow = 1; startRow <= selection.Rows.Count; startRow += chunkSize)
            {
                int endRow = Math.Min(startRow + chunkSize - 1, selection.Rows.Count);

                Excel.Range chunk = selection.Worksheet.Range[selection.Cells[startRow, 1], selection.Cells[endRow, selection.Columns.Count]];

                ProcessChunk(chunk);
            }
        }

        private void ProcessChunk(Excel.Range chunk)
        {
            foreach (Excel.Range cell in chunk)
            {
                if (cell.Value is string)
                {
                    cell.Value = cell.Value.Trim();
                }
            }
        }

        private void strongnoempty_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            foreach (Excel.Range cell in selection)
            {
                if (cell.Value is string)
                {
                    cell.Value = Regex.Replace(cell.Value, @"\s", "");
                }
            }
        }

        private void error20_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            int emptyCellCount = 0;
            foreach (Excel.Range row in selection.Rows)
            {
                foreach (Excel.Range cell in row.Cells)
                {
                    if (emptyCellCount >= 10000)
                    {
                        MessageBox.Show("Too Much Empty Cells", "10000+ emtpy cell found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (cell.HasFormula)
                    {
                        string originalFormula = cell.Formula;
                        string newFormula = "=IFERROR(" + originalFormula.Substring(1) + ",0)";
                        cell.Formula = newFormula;
                        emptyCellCount = 0;
                        continue;
                    }
                    emptyCellCount++;
                    continue;
                }
            }
        }

        private void error2empty_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            int emptyCellCount = 0;
            foreach (Excel.Range row in selection.Rows)
            {
                foreach (Excel.Range cell in row.Cells)
                {
                    if (emptyCellCount >= 10000)
                    {
                        MessageBox.Show("Too Much Empty Cells", "10000+ emtpy cell found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (cell.HasFormula)
                    {
                        string originalFormula = cell.Formula;
                        string newFormula = "=IFERROR(" + originalFormula.Substring(1) + ", \"\")";
                        cell.Formula = newFormula;
                        emptyCellCount = 0;
                        continue;
                    }
                    emptyCellCount++;
                    continue;
                }
            }
        }

        private void rm_bg_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            Excel.Range selection = (Excel.Range)ws.Application.Selection;
            selection.Interior.ColorIndex = Excel.Constants.xlNone;
            selection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }
    }
}
