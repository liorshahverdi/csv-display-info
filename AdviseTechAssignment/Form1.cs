using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AdviseTechAssignment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Listener for the 'Load Portfolio.csv file' button control
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            bool improper_file = false;

            openFileDialog1.Filter = "CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml";
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {
                    //boilerplate reference taken from http://csharp.net-informations.com/excel/csharp-read-excel.htm
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    Excel.Range range;

                    string str;
                    int rCnt = 0;
                    int cCnt = 0;

                    loadedPortfolios = new PortfolioCollection();

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    range = xlWorkSheet.UsedRange;

                    for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                    {
                        if (improper_file) break;
                        else
                        {
                            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                            {
                               str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                               string[] rowProperties = str.Split(',', '\t');

                               string realMarketValue = "";
                               for (int i = 2; i < rowProperties.Length; i++)
                               {
                                   realMarketValue += rowProperties[i];
                               }
                                   //ensure file has the proper attributes
                                   if (rCnt == 1)
                                   {
                                       if (rowProperties[0].Equals("Portfolio Name") &&
                                           rowProperties[1].Equals("Portfolio Code") &&
                                           rowProperties[2].Equals("Portfolio Market Value")) continue;
                                       else
                                       {
                                           improper_file = true;
                                           MessageBox.Show("The file you loaded does not contain Portfolios.\nPlease load the proper file");
                                           break;
                                       }
                                   }
                                //creates new portfolio instance from next entry in the file and adds it to our portfolio collection
                                Portfolio nextPortfolio = new Portfolio(rowProperties[0], rowProperties[1], realMarketValue);
                                loadedPortfolios.PortfolioEntities.Add(nextPortfolio);
                            }
                        }
                        
                    }

                    foreach (Portfolio p in loadedPortfolios.PortfolioEntities)
                    {
                        comboBox1.Items.Add(p.Name);
                        treeView1.Nodes.Add(p.Name);
                    }

                    if (!improper_file) MessageBox.Show("Loaded successfully!\n" + loadedPortfolios.PortfolioEntities.Count + " portfolios contained in " + file);

                    xlWorkBook.Close(true, null, null);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
                catch (IOException)
                {
                    MessageBox.Show(file + " did not load successfully.");
                }
            }
        }

        /// <summary>
        /// Listener for the 'Load PortfolioShareClass.csv file' button control 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            bool improper_file = false;

            openFileDialog2.Filter = "CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml";
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK)
            {
                string file = openFileDialog2.FileName;
                try
                {   //boilerplate reference taken from http://csharp.net-informations.com/excel/csharp-read-excel.htm
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    Excel.Range range;

                    string str;
                    int rCnt = 0;
                    int cCnt = 0;

                    loadedPortfolioShareClasses = new PortfolioShareClassCollection();

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    range = xlWorkSheet.UsedRange;

                    for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                    {
                        if (improper_file) break;
                        else
                        {
                            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                            {
                                str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                string[] rowProperties = str.Split(',', '\t');

                                //ensure file has the proper attributes
                                if (rCnt == 1)
                                {
                                    if (rowProperties[0].Equals("Portfolio Name") &&
                                        rowProperties[1].Equals("Portfolio Share Class Name") &&
                                        rowProperties[2].Equals("Portfolio Share Class Code") &&
                                        rowProperties[3].Equals("Portfolio Share Class Base Fee")) continue;
                                    else
                                    {
                                        improper_file = true;
                                        MessageBox.Show("The file you loaded does not contain Portfolio Share Classes.\nPlease load the proper file");
                                        break;
                                    }
                                }
                                //creates new portfolio share class instance from next entry in the file and adds it to our portfolio share class collection
                                PortfolioShareClass nextPortfolioShareClass = new PortfolioShareClass(rowProperties[0], rowProperties[1], rowProperties[2], rowProperties[3]);
                                loadedPortfolioShareClasses.PortfolioShareClassEntities.Add(nextPortfolioShareClass);
                            }
                        }
                    }

                    if (!improper_file) MessageBox.Show("Loaded successfully!\n" + loadedPortfolioShareClasses.PortfolioShareClassEntities.Count + " portfolio share classes contained in " + file);

                    xlWorkBook.Close(true, null, null);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
                catch (IOException)
                {
                    MessageBox.Show(file + " did not load successfully.");
                }
            }

        }

        //boilerplate reference taken from http://csharp.net-informations.com/excel/csharp-read-excel.htm
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// listener for the Portfolio ComboBox control
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loadedPortfolioShareClasses != null)
            {   //retrieve name selected in combobox
                string selected_name = comboBox1.SelectedItem.ToString();

                if (!selected_name.Equals("-")) dataGridView1.Visible = true; //make datagridview appear
                else dataGridView1.Visible = false; //hide datagridview

                treeView1.Visible = true;//make treeview appear

                List<PortfolioShareClass> retrievedList = null;
                retrievedList = loadedPortfolioShareClasses.PortfolioShareClassEntities;//list loaded from openfiledialog

                 IEnumerable<PortfolioShareClass> matched = from entry in retrievedList
                                                           where entry.ParentPortfolio == selected_name
                                                           select entry;

                 List<string[]> output = new List<string[]>();
                 foreach (PortfolioShareClass p in matched)
                 {
                    string[] properties = { p.ParentPortfolio, p.PortfolioShareClassName, p.PortfolioShareClassCode, p.PortfolioShareClassBaseFee };
                    output.Add(properties);
                 }
                 DataTable table = ConvertListToDataTable(output);
                 dataGridView1.DataSource = table;
            }
            else
            {
                MessageBox.Show("Portfolio Share Classes haven't been loaded yet!");
            }

            
        }

        
        //taken from http://www.dotnetperls.com/convert-list-datatable
        /// <summary>
        /// converts the List of strings to a DataTable object that can be displayed in our DataGridView control
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        static DataTable ConvertListToDataTable(List<string[]> list)
        {
            DataTable table = new DataTable();

            int columns = 0;
            foreach (var array in list)
            {
                if (array.Length > columns)
                {
                    columns = array.Length;
                }
            }
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add();
            }
            foreach (var array in list)
            {
                table.Rows.Add(array);
            }
            return table;
        }

        /// <summary>
        /// Listener for the Portfolio and PortfolioShareClass TreeView control
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (loadedPortfolioShareClasses != null)
            {
                foreach (TreeNode n in treeView1.Nodes)
                {
                    string root_name = n.Text;
                    List<PortfolioShareClass> retrievedList = null;
                    retrievedList = loadedPortfolioShareClasses.PortfolioShareClassEntities;//list loaded from openfiledialog
                    
                    //populate a new collection of portfolio share class instances that have the same name as the root portfolio
                    IEnumerable<PortfolioShareClass> matched = from entry in retrievedList
                                                               where entry.ParentPortfolio == root_name
                                                               select entry;

                    List<string> children = new List<string>();
                    if (n.Nodes.Count == 0)
                    {
                        foreach (PortfolioShareClass psc in matched)
                        {
                            //add portfolio share class instance to collection of child nodes for this portfolio
                            string share_class_name = psc.PortfolioShareClassName;
                            n.Nodes.Add(share_class_name);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Portfolio Share Classes haven't been loaded yet!");
            }
        }

        /// <summary>
        /// terminates the application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double sum = 0;
            try
            {
                List<Portfolio> csc = loadedPortfolios.PortfolioEntities;
                foreach (Portfolio c in csc)
                {
                    char[] getRidOfThese = { '"' };
                    string next = c.MarketValue.Substring(3).TrimEnd(getRidOfThese);
                    double realNumber = Convert.ToDouble(next);
                    sum += realNumber;
                }
                MessageBox.Show("sum of portfolio market values = " + sum);
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Portfolio.csv has not been loaded yet!");
            }
        }
    }
}