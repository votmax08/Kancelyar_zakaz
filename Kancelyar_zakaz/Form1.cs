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
using System.Xml.Serialization;

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
            if (dataGridView1.Rows.Count > 0)
                if (dataGridView1.SelectedCells.Count > 0)
                    button2.Visible = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Width = 0;
            foreach (DataGridViewColumn col in dataGridView1.Columns)
                dataGridView1.Width += col.Width;
            string path = @"Data.xml";
            XmlSerializer formatter = new XmlSerializer(typeof(Predmet[]));
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                Predmet[] dePredmet = (Predmet[])formatter.Deserialize(fs);
                for (int i = 0; i < dePredmet.Length; i++)
                {
                    NewPredmet p = new NewPredmet(dePredmet[i].Id, dePredmet[i].Name, dePredmet[i].Art);
                    comboBox1.Items.Add(p.Name);
                    predmets.Add(p);
                }
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int iSelected = comboBox1.SelectedIndex;
            if (iSelected>=0)
                if (!predmets[iSelected].isAdded)
                {
                    decimal N = numericUpDown1.Value;
                    dataGridView1.Rows.Add(1);
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        switch (i)
                        {
                            case 0:
                                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value = predmets[iSelected].Name;
                                break;
                            case 1:
                                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value = predmets[iSelected].Art;
                                break;
                            case 2:
                                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value = N.ToString();
                                break;
                        }
                    Focus();
                    if (dataGridView1.Rows.Count > 0)
                        if (dataGridView1.SelectedCells.Count > 0)
                            button2.Visible = true;
                }
                else
                {
                    //MessageBox
                    MessageBox.Show("Элемент уже добавлен!");
                }
        }

        List<NewPredmet> predmets = new List<NewPredmet>();

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
            {
                dataGridView1.Rows.RemoveAt(cell.RowIndex);
            }
            if (dataGridView1.SelectedCells.Count == 0)
                button2.Visible = false;
        }
    }
    [Serializable]
    public class Predmet
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Art { get; set; }

        public Predmet()
        {
        }
        public Predmet(string id, string name, string art)
        {
            Id = id;
            Name = name;
            Art = art;
        }
    }
    public class NewPredmet: Predmet
    {
        public string KomusArt { get; set; }
        public bool isAdded { get; set; }
        public NewPredmet()
        {
        }
        public NewPredmet(string id, string name, string art, string komusArt="", bool added=false)
            :base(id, name, art)
        {
            KomusArt = komusArt;
            isAdded = added;
        }
    }
}
