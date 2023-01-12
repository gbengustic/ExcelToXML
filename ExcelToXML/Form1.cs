using ClosedXML.Excel;
using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace ExcelToXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            TxtOutputPath.Text = Properties.Settings.Default.TxtOutputPath;
            TxtInputExcelPath.Text = Properties.Settings.Default.TxtInputExcelPath;
            TxtPymatFile.Text = Properties.Settings.Default.TxtPymatFile;
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(TxtOutputPath.Text))
            {
                MessageBox.Show("Please select an output path.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (string.IsNullOrEmpty(TxtPymatFile.Text))
            {
                MessageBox.Show("Please select an Pymat file.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            BtnStart.Enabled = false;
            BtnExcelToPymat.Enabled = false;
            backgroundWorker2.RunWorkerAsync();
        }
        private void BtnBrowsePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                TxtOutputPath.Text = folderBrowserDialog.SelectedPath;
                Properties.Settings.Default.TxtOutputPath = TxtOutputPath.Text;
                Properties.Settings.Default.Save();
            }
        }
        public static DataTable ImportExceltoDatatable(string filePath)
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        if (row.FirstCellUsed() == null) continue;
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dt;
            }
        }

        private void BtnBrowseInputFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                TxtInputExcelPath.Text = openFileDialog.FileName;
                Properties.Settings.Default.TxtInputExcelPath = TxtInputExcelPath.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            try
            {
                DataTable dataTable = ImportExceltoDatatable(TxtInputExcelPath.Text);
                string asset = File.ReadAllText("asset.txt");
                string bumpmap_asset = File.ReadAllText("bumpmap_asset.txt");
                string input = File.ReadAllText("input.xml");
                string material = File.ReadAllText("material.txt");
                string texture_asset = File.ReadAllText("texture_asset.txt");
                int texture_bumpmap_asset_counter = 0;
                StringBuilder materialString = new StringBuilder();
                StringBuilder assetString = new StringBuilder();

                foreach (DataRow dataRow in dataTable.Rows)
                {
                    string tempAsset = "";
                    string TempMaterial = material.Replace(":::A:::", dataRow[0].ToString())
                        .Replace(":::MAT_ID:::", (dataTable.Rows.IndexOf(dataRow) + 1).ToString())
                        .Replace(":::Q:::", string.IsNullOrEmpty(dataRow[16].ToString().Trim()) ? 0.ToString() : dataRow[16].ToString())
                        .Replace(":::P:::", string.IsNullOrEmpty(dataRow[15].ToString().Trim()) ? 0.ToString() : dataRow[15].ToString())
                        .Replace(":::O:::", string.IsNullOrEmpty(dataRow[14].ToString().Trim()) ? 0.ToString() : dataRow[14].ToString())
                        .Replace(":::R:::", string.IsNullOrEmpty(dataRow[17].ToString().Trim()) ? 0.ToString() : dataRow[17].ToString())
                        .Replace(":::B:::", string.IsNullOrEmpty(dataRow[1].ToString().Trim()) ? 0.ToString() : dataRow[1].ToString())
                        .Replace(":::C:::", string.IsNullOrEmpty(dataRow[2].ToString().Trim()) ? 0.ToString() : dataRow[2].ToString())
                        .Replace(":::D:::", string.IsNullOrEmpty(dataRow[3].ToString().Trim()) ? 0.ToString() : dataRow[3].ToString())
                        .Replace(":::E:::", string.IsNullOrEmpty(dataRow[4].ToString().Trim()) ? 0.ToString() : dataRow[4].ToString())
                        .Replace(":::F:::", string.IsNullOrEmpty(dataRow[5].ToString().Trim()) ? 0.ToString() : dataRow[5].ToString())
                        .Replace(":::G:::", string.IsNullOrEmpty(dataRow[6].ToString().Trim()) ? 0.ToString() : dataRow[6].ToString())
                        .Replace(":::I:::", string.IsNullOrEmpty(dataRow[8].ToString().Trim()) ? 0.ToString() : dataRow[8].ToString())
                        .Replace(":::J:::", string.IsNullOrEmpty(dataRow[9].ToString().Trim()) ? 0.ToString() : dataRow[9].ToString())
                        .Replace(":::K:::", string.IsNullOrEmpty(dataRow[10].ToString().Trim()) ? 0.ToString() : dataRow[10].ToString())
                        .Replace(":::L:::", string.IsNullOrEmpty(dataRow[11].ToString().Trim()) ? 0.ToString() : dataRow[11].ToString())
                        .Replace(":::N:::", string.IsNullOrEmpty(dataRow[13].ToString().Trim()) ? 0.ToString() : dataRow[13].ToString())

                        .Replace(":::S:::", string.IsNullOrEmpty(dataRow[18].ToString().Trim()) ? 0.ToString() : dataRow[18].ToString())
                        .Replace(":::T:::", string.IsNullOrEmpty(dataRow[19].ToString().Trim()) ? 0.ToString() : dataRow[19].ToString())
                        .Replace(":::U:::", string.IsNullOrEmpty(dataRow[20].ToString().Trim()) ? 0.ToString() : dataRow[20].ToString())
                        .Replace(":::V:::", string.IsNullOrEmpty(dataRow[21].ToString().Trim()) ? 0.ToString() : dataRow[21].ToString())
                        .Replace(":::W:::", string.IsNullOrEmpty(dataRow[22].ToString().Trim()) ? 0.ToString() : dataRow[22].ToString())
                        .Replace(":::X:::", string.IsNullOrEmpty(dataRow[23].ToString().Trim()) ? 0.ToString() : dataRow[23].ToString())
                        .Replace(":::Y:::", string.IsNullOrEmpty(dataRow[24].ToString().Trim()) ? 0.ToString() : dataRow[24].ToString())
                        .Replace(":::Z:::", string.IsNullOrEmpty(dataRow[25].ToString().Trim()) ? 0.ToString() : dataRow[25].ToString())

                        ;


                    if (string.IsNullOrEmpty(dataRow[7].ToString().Trim()))
                    {
                        TempMaterial = TempMaterial.Replace(":::a:texture-asset:::", "");
                    }
                    else
                    {
                        texture_bumpmap_asset_counter++;
                        string Temptexture_asset = texture_asset.Replace(":::ID:::", texture_bumpmap_asset_counter.ToString());
                        TempMaterial = TempMaterial.Replace(":::a:texture-asset:::", Temptexture_asset);
                        tempAsset = asset.Replace(":::ID:::", texture_bumpmap_asset_counter.ToString()).Replace(":::URL:::", string.IsNullOrEmpty(dataRow[7].ToString().Trim()) ? 0.ToString() : dataRow[7].ToString());
                        assetString.Append(tempAsset);
                        assetString.AppendLine();
                    }

                    if (string.IsNullOrEmpty(dataRow[12].ToString().Trim()))
                    {
                        TempMaterial = TempMaterial.Replace(":::a:bumpmap-asset:::", "");
                    }
                    else
                    {
                        texture_bumpmap_asset_counter++;
                        string Tempbumpmap_asset = bumpmap_asset.Replace(":::ID:::", texture_bumpmap_asset_counter.ToString());
                        TempMaterial = TempMaterial.Replace(":::a:bumpmap-asset:::", Tempbumpmap_asset);
                        tempAsset = asset.Replace(":::ID:::", texture_bumpmap_asset_counter.ToString()).Replace(":::URL:::", string.IsNullOrEmpty(dataRow[12].ToString().Trim()) ? 0.ToString() : dataRow[12].ToString());
                        assetString.Append(tempAsset);
                        assetString.AppendLine();
                    }
                    materialString.Append(TempMaterial);
                    materialString.AppendLine();
                }
                materialString.ToString();
                input = input.Replace(":::MATERIAL:::", materialString.ToString()).Replace(":::a:attributes:::", assetString.ToString());
                File.WriteAllText(Path.Combine(TxtOutputPath.Text, "result.pymat"), input);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BtnStart.Enabled = true;
            BtnExcelToPymat.Enabled = true;
            MessageBox.Show("Processing complete", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnSelectPymatFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Pymat Files|*.pymat";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                TxtPymatFile.Text = openFileDialog.FileName;
                Properties.Settings.Default.TxtPymatFile = TxtPymatFile.Text;
                Properties.Settings.Default.Save();
            }
        }

        private void BtnExcelToPymat_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(TxtOutputPath.Text))
            {
                MessageBox.Show("Please select an output path.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (string.IsNullOrEmpty(TxtInputExcelPath.Text))
            {
                MessageBox.Show("Please select an Excel file.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            BtnStart.Enabled = false;
            BtnExcelToPymat.Enabled = false;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            try
            {
                var workbook = new XLWorkbook("Template.xlsx");
                var ws = workbook.Worksheet(1);
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(Path.Combine(TxtPymatFile.Text));
                int row = 2;
                foreach (XmlNode node in xDoc.LastChild.ChildNodes[2])
                {
                    string A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z;
                    A = node.Attributes[0].Value;
                    ws.Cell(row, 1).Value = A;
                    foreach (XmlNode node1 in node.ChildNodes[0].ChildNodes)
                    {

                        if (node1.Name == "a:color-srgb")
                        {
                            B = node1.ChildNodes[0].InnerText;
                            C = node1.ChildNodes[1].InnerText;
                            D = node1.ChildNodes[2].InnerText;
                            ws.Cell(row, 2).Value = B;
                            ws.Cell(row, 3).Value = C;
                            ws.Cell(row, 4).Value = D;
                        }
                        else if (node1.Name == "a:diffuse")
                        {
                            E = node1.ChildNodes[0].InnerText;
                            ws.Cell(row, 5).Value = E;
                        }
                        else if (node1.Name == "a:reflective")
                        {
                            F = node1.ChildNodes[0].InnerText;
                            ws.Cell(row, 6).Value = F;
                        }
                        else if (node1.Name == "a:transparent")
                        {
                            G = node1.ChildNodes[0].InnerText;
                            ws.Cell(row, 7).Value = G;
                        }
                        else if (node1.Name == "a:texture-asset")
                        {
                            H = node1.ChildNodes[1].InnerText;
                            ws.Cell(row, 8).Value = xDoc.LastChild.LastChild.ChildNodes[Convert.ToInt32(H) - 1].ChildNodes[0].Attributes[0].InnerText;
                        }
                        else if (node1.Name == "a:texture-mapping")
                        {
                            I = node1.ChildNodes[0].InnerText;
                            ws.Cell(row, 9).Value = I;
                        }
                        else if (node1.Name == "a:texture-multiply-color")
                        {
                            J = node1.InnerText;
                            ws.Cell(row, 10).Value = J;
                        }
                        else if (node1.Name == "a:texture-size")
                        {
                            K = node1.ChildNodes[0].InnerText;
                            L = node1.ChildNodes[1].InnerText;
                            ws.Cell(row, 11).Value = K;
                            ws.Cell(row, 12).Value = L;
                        }
                        else if (node1.Name == "a:bumpmap-asset")
                        {
                            M = node1.ChildNodes[1].InnerText;
                            ws.Cell(row, 13).Value = xDoc.LastChild.LastChild.ChildNodes[Convert.ToInt32(M) - 1].ChildNodes[0].Attributes[0].InnerText;
                        }
                        else if (node1.Name == "a:bumpmap-angle")
                        {
                            N = node1.InnerText;
                            ws.Cell(row, 14).Value = N;
                        }
                        else if (node1.Name == "a:construction-type")
                        {
                            O = node1.InnerText;
                            ws.Cell(row, 15).Value = O;
                        }
                        else if (node1.Name == "a:core-material")
                        {
                            P = node1.InnerText;
                            ws.Cell(row, 16).Value = P;
                        }
                        else if (node1.Name == "a:article-no")
                        {
                            Q = node1.InnerText;
                            ws.Cell(row, 17).Value = Q;
                        }
                        else if (node1.Name == "a:description")
                        {
                            R = node1.InnerText;
                            ws.Cell(row, 18).Value = R;
                        }
                        else if (node1.Name == "a:attribute" && node1.Attributes[1].InnerText== "Weight,spec.(g/cm³)")
                        {
                            S = node1.InnerText;
                            ws.Cell(row, 19).Value = S;
                        }
                        else if (node1.Name == "a:attribute" && node1.Attributes[1].InnerText == "Price")
                        {
                            T = node1.InnerText;
                            ws.Cell(row, 20).Value = T;
                        }
                        else if (node1.Name == "a:user-attri1")
                        {
                            U = node1.InnerText;
                            ws.Cell(row, 21).Value = U;
                        }
                        else if (node1.Name == "a:user-attri2")
                        {
                            V = node1.InnerText;
                            ws.Cell(row, 22).Value = V;
                        }
                        else if (node1.Name == "a:user-attri3")
                        {
                            W = node1.InnerText;
                            ws.Cell(row, 23).Value = W;
                        }
                        else if (node1.Name == "a:user-attri4")
                        {
                            X = node1.InnerText;
                            ws.Cell(row, 24).Value = X;
                        }
                        else if (node1.Name == "a:user-attri5")
                        {
                            Y = node1.InnerText;
                            ws.Cell(row, 25).Value = Y;
                        }
                        else if (node1.Name == "a:user-attri6")
                        {
                            Z = node1.InnerText;
                            ws.Cell(row, 26).Value = Z;
                        }
                    }
                    row++;
                }
                workbook.SaveAs(Path.Combine(TxtOutputPath.Text, "Output.xlsx"));
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BtnStart.Enabled = true;
            BtnExcelToPymat.Enabled = true;
            MessageBox.Show("Processing complete", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
