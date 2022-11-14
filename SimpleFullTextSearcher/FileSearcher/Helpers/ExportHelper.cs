using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleFullTextSearcher.FileSearcher.Helpers
{
    internal static class ExportHelper
    {
        public static bool ExportTreeViewNodesToExcel(TreeView treeView, FileInfo fileInfo = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();
            var worksheet = excel.Workbook.Worksheets.Add("TreeView Export");
            int rowCounter = 0;

            RecurseNodes(treeView.Nodes, 1);

            void RecurseNodes(TreeNodeCollection currentNode, int col)
            {
                foreach (TreeNode node in currentNode)
                {
                    rowCounter = rowCounter + 1;
                    worksheet.Cells[rowCounter, col].Value = node.Text;
                    if (node.FirstNode != null)
                        RecurseNodes(node.Nodes, col + 1);
                }
            }

            if (fileInfo is null)
            {
                var fileDialogue = new SaveFileDialog();
                fileDialogue.Filter = "xlsx Dateien (*.xlsx)|*.xlsx";
                if (fileDialogue.ShowDialog() == DialogResult.OK)
                    fileInfo = new FileInfo(fileDialogue.FileName);
            }
            excel.SaveAs(fileInfo);
            return fileInfo.Exists;
        }
    }
}
