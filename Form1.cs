using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace version04
{
    public partial class Form1 : Form
    {
        MolotdelDataSetTableAdapters.QueriesTableAdapter adapt = new MolotdelDataSetTableAdapters.QueriesTableAdapter();
        public Form1()
        {
            InitializeComponent();
        }

        private void glavnayaBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.glavnayaBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.molotdelDataSet);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "molotdelDataSet.Ychastnik". При необходимости она может быть перемещена или удалена.
            this.ychastnikTableAdapter.Fill(this.molotdelDataSet.Ychastnik);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "molotdelDataSet.Calendar". При необходимости она может быть перемещена или удалена.
            this.calendarTableAdapter.Fill(this.molotdelDataSet.Calendar);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "molotdelDataSet.Otryadi". При необходимости она может быть перемещена или удалена.
            this.otryadiTableAdapter.Fill(this.molotdelDataSet.Otryadi);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "molotdelDataSet.Goroda". При необходимости она может быть перемещена или удалена.
            this.gorodaTableAdapter.Fill(this.molotdelDataSet.Goroda);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "molotdelDataSet.Sotrudnik". При необходимости она может быть перемещена или удалена.
            this.sotrudnikTableAdapter.Fill(this.molotdelDataSet.Sotrudnik);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "molotdelDataSet.Meropriyatii". При необходимости она может быть перемещена или удалена.
            this.meropriyatiiTableAdapter.Fill(this.molotdelDataSet.Meropriyatii);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "molotdelDataSet.Glavnaya". При необходимости она может быть перемещена или удалена.
            this.glavnayaTableAdapter.Fill(this.molotdelDataSet.Glavnaya);
            adapt = new MolotdelDataSetTableAdapters.QueriesTableAdapter();
        }

        private void btDelete_Click(object sender, EventArgs e)
        {
            adapt.DelMeropriyatii(int.Parse(tBidMeropriyatii.Text));
            this.meropriyatiiTableAdapter.Fill(this.molotdelDataSet.Meropriyatii);
        }
        //поиск прошедших мероприятий
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                //this.printPeopleTableAdapter.Fill(this.verTwoDataSet.printPeople, peremenToolStripTextBox.Text);
                this.printOldMeropiyatiiTableAdapter.Fill(this.molotdelDataSet.printOldMeropiyatii, peremen:dateTimePicker3.Value);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        //поиск будущих мероприятий
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                this.printNewMeropiyatiiTableAdapter.Fill(this.molotdelDataSet.printNewMeropiyatii, peremen: dateTimePicker3.Value);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        //вывод в Excel
        public void button1_Click(object sender, EventArgs e)
        {
            try { 
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ExWrBook;
                Microsoft.Office.Interop.Excel.Worksheet ExWrSheet;
                ExWrBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                ExWrSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExWrBook.Worksheets.get_Item(1);
                ExcelApp.Visible = true;

                //ExcelApp.Application.Workbooks.Add(Type.Missing);//создать рабочую книгу

            ExcelApp.Cells[1, 1] = "Номер мероприятия";
            ExcelApp.Cells[1, 2] = "Дата проведения";
            ExcelApp.Cells[1, 3] = "Наименование";
            ExcelApp.Cells[1, 4] = "Содержание";
            ExcelApp.Cells[1, 5] = "Организатор";
            ExcelApp.Cells[1, 6] = "Итоги";
            ExcelApp.Cells[1, 7] = "Событие";
            ExcelApp.Cells[1, 8] = "Количество участников";

            for (int i=0; i < pMView.RowCount; i++)
            {
                ExcelApp.Cells[ i+2 , 1] = pMView.Rows[i].Cells[0].Value.ToString();
                    ExcelApp.Cells[i + 2, 2] = pMView.Rows[i].Cells[1].Value.ToString();
                    ExcelApp.Cells[i + 2, 3] = pMView.Rows[i].Cells[2].Value.ToString();
                    ExcelApp.Cells[i + 2, 4] = pMView.Rows[i].Cells[3].Value.ToString();
                    ExcelApp.Cells[i + 2, 5] = pMView.Rows[i].Cells[4].Value.ToString();
                    ExcelApp.Cells[i + 2, 6] = pMView.Rows[i].Cells[5].Value.ToString();
                    ExcelApp.Cells[i + 2, 7] = pMView.Rows[i].Cells[6].Value.ToString();
                    ExcelApp.Cells[i + 2, 8] = pMView.Rows[i].Cells[7].Value.ToString();
                }
                ExcelApp.UserControl = true;
            }
            catch (Exception error)
            {
                MessageBox.Show("Экспорт " + error.Source + " in MS Excel. Успешно!");
                return;
            }
        }

        private void btDobavit_Click(object sender, EventArgs e)
        {
            adapt.inMeropriyatii(
                int.Parse(tBidMeropriyatii.Text),
                DateTime.Parse(dateTimePicker1.Text),
                tbNazvanie.Text,
                tbOpisanie.Text,
                Convert.ToInt32(comboBox1Organizator.SelectedValue),
                tbItogi.Text,
                Convert.ToInt32(comboBox5.SelectedValue),
                int.Parse(tbKol_voYshastnikov.Text)
                );
            this.meropriyatiiTableAdapter.Fill(this.molotdelDataSet.Meropriyatii);
        }

        private void btObnovit_Click(object sender, EventArgs e)
        {
            adapt.upMeropriyatii(
                int.Parse(tBidMeropriyatii.Text),
                DateTime.Parse(dateTimePicker1.Text),
                tbNazvanie.Text,
                tbOpisanie.Text,
                Convert.ToInt32(comboBox1Organizator.SelectedValue),
                tbItogi.Text,
                Convert.ToInt32(comboBox5.SelectedValue),
                int.Parse(tbKol_voYshastnikov.Text)
                );
            this.meropriyatiiTableAdapter.Fill(this.molotdelDataSet.Meropriyatii);
        }
        //мероприятия за неделю или опр промежуток
        private void button4_Click(object sender, EventArgs e)
        {
            this.printNedMerpTableAdapter.Fill(this.molotdelDataSet.PrintNedMerp, dateTimePicker4.Value, dateTimePicker5.Value);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.printOldMeropiyatiiTableAdapter.Fill(this.molotdelDataSet.printOldMeropiyatii, peremen: dateTimePicker3.Value);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}
