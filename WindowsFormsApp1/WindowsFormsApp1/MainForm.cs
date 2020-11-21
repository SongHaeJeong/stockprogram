using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        private string filePath = "";

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "File Name";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            if(OFD.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                textBox1.Text = OFD.FileName;
                filePath = OFD.FileName;
            }

            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (filePath != "")
            {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = application.Workbooks.Open(Filename: @filePath);
                Worksheet worksheet1 = workbook.Worksheets.get_Item("TEST");
                application.Visible = false;
                Range range = worksheet1.UsedRange;
                String data = "";
                data += "\n";
                long price = 0; // 가격
                long type = 0; //매수, 매도의 수량
                long totalSell = 0; //매도 총액
                long totalBuy = 0; // 매수 총액                
                long purePrice = 0;
                long rowCount = range.Rows.Count;
                for (int i = 2; i <= rowCount; ++i)
                {

                    price = (long)(range.Cells[i, 2]).Value2;
                    type = (long)(range.Cells[i, 6]).Value2;
                    if(type < 0) //매도
                    {
                        long sellPrice = price * type;
                        totalSell += sellPrice;
                    }
                    else // 매수
                    {
                        long buyPrice = price * type;
                        totalBuy += buyPrice;
                    }
                    
                }

                data += "거래 횟수 : " + (range.Rows.Count - 1) + "\n\n";
                data += "종가 : " + (range.Cells[2, 2]).Value2 + "\n\n";
                data += "매수 : " + (totalBuy) + "\n\n";
                data += "매도 : " + (totalSell) + "\n\n";
                //순매수 계산
                purePrice = totalBuy + totalSell;
                data += "순매수 : " + (purePrice) + "\n";

                richTextBox1.Text = data;

                DeleteObject(worksheet1);
                DeleteObject(workbook);
                application.Quit();
                DeleteObject(application);


            }
        }
       
        private void DeleteObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("메모리 할당을 해제하는 중 문제가 발생하였습니다." + ex.ToString(), "경고!");
            }
            finally
            {
                GC.Collect();
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
