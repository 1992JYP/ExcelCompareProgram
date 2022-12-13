using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLoad
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            string[] data1 = new string[5];
            string[] data2 = new string[3] { "0", "0", "0" };




            for (int i = 0; i < 10; i++)
            {
                if (i < 3)
                {
                    data1[0] = (i + 1).ToString();
                    switch (i)
                    {
                        case 0:
                            data1[1] = "0";
                            data1[2] = "0";
                            data1[3] = "OFF";
                            data1[4] = "0";
                            break;
                        case 1:
                            data1[1] = "8";
                            data1[2] = "0";
                            data1[3] = "ON";
                            data1[4] = "10";
                            break;
                        case 2:
                            data1[1] = "20";
                            data1[2] = "0";
                            data1[3] = "OFF";
                            data1[4] = "0";
                            break;
                    }
                    dataGridView1.Rows.Add(data1);
                }
                else
                {
                    data2[0] = (i+1).ToString();
                    dataGridView1.Rows.Add(data2);
                }
            }

            string[] data3 = new string[15];

            for (int i = 0; i < 15; i++)
            {
                data3[i]=i.ToString();
            }

            dataGridView1.Rows.Add (data3);

        }



        private void button1_Click(object sender, EventArgs e)
        {
            //Excel.Application excelApp = null;
            //Excel.Workbook workbook = null;
            //Excel.Worksheet worksheet = null;  
            //Excel.Worksheet worksheet2 = null;
        }
    }
}
