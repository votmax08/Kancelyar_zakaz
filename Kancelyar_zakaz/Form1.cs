using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel =  Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;


namespace Kancelyar_zakaz
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        List<Item> items = new List<Item>();
        private static object[,] loadCellByCell(int row, int maxColNum, _Worksheet osheet)
        {
            var list = new object[2, maxColNum + 1];
            for (int i = 1; i <= maxColNum; i++)
            {
                var RealExcelRangeLoc = osheet.Range[(object)osheet.Cells[row, i], (object)osheet.Cells[row, i]];
                object valarrCheck;
                try
                {
                    valarrCheck = RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];
                }
                catch
                {
                    valarrCheck = (object)RealExcelRangeLoc.Value2;
                }
                list[1, i] = valarrCheck;
            }
            return list;
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

            Excel.Application ExcelObj = null;
            _Workbook ecelbook = null;
            try
            {
                ExcelObj = new Excel.Application();
                ExcelObj.DisplayAlerts = false;
                const string f = @"\\Mac\Home\Desktop\Kancelyar_zakaz\Items.xlsx";
                ecelbook = ExcelObj.Workbooks.Open(f, 0, true, 5, "", "", false, XlPlatform.xlWindows);
                var sheets = ecelbook.Sheets;
                var maxNumSheet = sheets.Count;
                for (int i = 1; i <= maxNumSheet; i++)
                {
                    var osheet = (_Worksheet)ecelbook.Sheets[i];
                    Range excelRange = osheet.UsedRange;

                    int maxColNum;
                    int lastRow;
                    try
                    {
                        maxColNum = excelRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                        lastRow = excelRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                    }
                    catch
                    {
                        maxColNum = excelRange.Columns.Count;
                        lastRow = excelRange.Rows.Count;
                    }

                    object[,] nameColArr = null;
                    for (int l = 1; l <= lastRow; l++)
                    {
                        Range RealExcelRangeLoc = osheet.Range[(object)osheet.Cells[l, 1], (object)osheet.Cells[l, maxColNum]];
                        object[,] valarr = null;
                        try
                        {
                            var valarrCheck = RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];
                            if (valarrCheck is object[,] || valarrCheck == null)
                                valarr = (object[,])RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];
                        }
                        catch
                        {
                            valarr = loadCellByCell(l, maxColNum, osheet);
                        }
                        if (l == 1)
                            nameColArr = valarr;
                        else
                        {
                            Item tempItem = new Item();
                            for (int j = 1; j <= nameColArr.Length; j++)
                                switch (nameColArr[1, j].ToString())
                                {
                                    case "id":
                                        tempItem.Id = Convert.ToInt32(valarr[1, j].ToString());
                                        break;
                                    case "name":
                                        tempItem.Name = valarr[1, j].ToString();
                                        break;
                                    case "code":
                                        if (valarr[1, j] != null)
                                            tempItem.Code = valarr[1, j].ToString();
                                        else
                                            tempItem.Code = null;
                                        break;
                                    case "unit":
                                        tempItem.Unit = valarr[1, j].ToString();
                                        break;
                                    case "category":
                                        if (valarr[1, j] != null)
                                            tempItem.Category = valarr[1, j].ToString();
                                        else
                                            tempItem.Category = "";
                                        break;
                                    case "komusart":
                                        if (valarr[1, j] != null)
                                            tempItem.KomusArt = valarr[1, j].ToString();
                                        else
                                            tempItem.KomusArt = "";
                                        break;
                                }
                            if (tempItem.Category == "")
                                tempItem.Category = tempItem.CorrectName();
                            items.Add(tempItem);
                        }
                    }
                }
            }
            finally
            {
                if (ecelbook != null)
                {
                    ecelbook.Close();
                    Marshal.ReleaseComObject(ecelbook);
                }
                if (ExcelObj != null) ExcelObj.Quit();
            }
            for (int i=0; i < items.Count; i++)
            {
                if (!comboBox1.Items.Contains(items[i].Category))
                    comboBox1.Items.Add(items[i].Category);
                //comboBox2.Items.Add(items[i].Name);
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int iSelected = comboBox1.SelectedIndex;

            if (iSelected >= 0)
            {

                //if (!predmets[iSelected].isOrdered)//isAdded function to find this element

                //{

                decimal N = numericUpDown1.Value;

                dataGridView1.Rows.Add(comboBox1.Text, comboBox2.Text, N.ToString(), textBox1.Text);

                //dataGridView1.Rows.Add(1);

                //for (int i = 1; i <= dataGridView1.Columns.Count; i++)

                //    switch (i)

                //    {

                //        case 0:



                //            dataGridView1.Rows[dataGridView1.Rows.Count].Cells[i].Value = predmets[iSelected].Name;

                //            break;

                //        case 1:

                //            dataGridView1.Rows[dataGridView1.Rows.Count].Cells[i].Value = predmets[iSelected].Articul;

                //            break;

                //        case 2:

                //            dataGridView1.Rows[dataGridView1.Rows.Count].Cells[i].Value = N.ToString();

                //            break;

                //    }

                Focus();

                if (dataGridView1.Rows.Count > 0)

                    if (dataGridView1.SelectedCells.Count > 0)

                        button2.Visible = true;

            }

            //}

            //else

            //{

            //MessageBox

            //MessageBox.Show("Элемент уже добавлен!");

            //}



            //set Form to default



            label3.Visible = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        List<Tuple<string, int>> indexList = new List<Tuple<string, int>>();
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Enabled = true;
            comboBox2.Items.Clear();
            indexList.Clear();

            for (int i = 0; i < items.Count; i++)
                if (comboBox1.Items[comboBox1.SelectedIndex].ToString() == items[i].Category)
                {
                    comboBox2.Items.Add(items[i].CorrectName());
                    Tuple<string, int> tuple = new Tuple<string, int>(items[i].CorrectName(), i);
                    indexList.Add(tuple);
                }

            numericUpDown1.Enabled = false;
            checkBox1.Enabled = false;
            label3.Visible = false;
            //toolStripStatusLabel1.Text = comboBox1.SelectedIndex.ToString();
        }

        Item selectedItem = new Item();
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown1.Enabled = true;
            checkBox1.Enabled = true;
            label3.Visible = true;

            selectedItem = null;
            string selectedName = comboBox2.Text;
            for (int i=0; i<indexList.Count; i++)
                if (indexList[i].Item1 == selectedName)
                    selectedItem = items[indexList[i].Item2];

            label3.Text = selectedItem.Unit;
            textBox1.Text = selectedItem.KomusArt;
        }

        private void SelectionList(string sCategory, object sender)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.Enabled = checkBox1.Checked;
        }

        //Unicum categorys of Item list ???Items globalen
        private List<string> ListOfCategorys(List<Item> items)
        {
            List<string> list = new List<string>();
            for (int i = 0; i < items.Count; i++)
                if (!list.Contains(items[i].Category))
                    list.Add(items[i].Category);
            return list;
        }
        //Find element in list ???Items globalen
        //private int FindElement(List<Item> items, string sVar, string ElementType, int FirstIndex = 0)
        //{
        //    int result = -1;
        //    for (int i = 0; i < items.Count; i++)
        //    {
        //        switch (ElementType)
        //        {
        //            case "Category":

        //        }
        //    }
        //}

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            textBox1.Enabled = !textBox1.Enabled;
            if (textBox1.Enabled)
            {
                textBox1.Focus();
                textBox1.SelectAll();
            }
        }
    }
    //[Serializable]
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public string Unit { get; set; }
        public string KomusArt { get; set; }
        public string Category { get; set; }
        public Item(int id, string name, string code, string unit, string komus, string cat)
        {
            Id = id;
            Name = name;
            Code = code;
            Unit = unit;
            KomusArt = komus;
            Category = cat;
        }
        public Item()
        {

        }
        public string CorrectName()
        {
            string outputtext = "";
            for (int i = 0; i < Name.Length; i++)
            {
                if (Name[i] == '_')//switch if other uncorrect chars
                    outputtext += " ";
                else
                    outputtext += Name[i].ToString();
            }
            return outputtext;
        }
    }

    public class OrderItem : Item
    {
        public bool isOrdered { get; set; }
        public string UserName { get; set; }
        public string OrderTime { get; set; }
        public int Count { get; set; }
        public OrderItem(int id, string name, string code, string unit, string komus, string cat, string user, string time, int count)
            : base(id, name, code, unit, komus, cat)
        {
            UserName = user;
            OrderTime = time;
            Count = count;
            isOrdered = false;
        }
    }

}