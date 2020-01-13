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
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;
using Ionic.Zip;
using System.Data.SQLite;
using System.Configuration;
using System.Threading;
using System.Globalization;
namespace WindowsFormsApp3
{

    public partial class Form1 : Form
    {



        public Form1()
        {

            InitializeComponent();

            //
            button1.Visible = true;
            //
            panel3.Visible = false;

            pnl_print.Visible = false;
            pnl_password.Visible = false;
            //
            pnl_menu.Visible = false;
            panel1.Visible = false;
            panel2.Visible = false;
            //
            button4.Visible = true;
            button6.Visible = true;
            //

        }

        //Environment.CurrentDirectory
        SQLiteConnection con = new SQLiteConnection(@"Data Source=database.db;Version=3;New=False;Compress=True;");

        //  SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=E:\michael\downloads\center\WindowsFormsApp3\WindowsFormsApp3\Database1.mdf;Initial Catalog=test;Integrated Security=True");
        /*  <connectionStrings>
  <add name="WindowsFormsApp3.Properties.Settings.BabakConnectionString"
      connectionString="Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database1.mdf;Initial Catalog=test;Integrated Security=True"
      providerName="System.Data.SqlClient" />
</connectionStrings>*/
        int exitt = 0;
    
        int select_index = 0;
        int index_cashier = 0;
        int tottal = 0;
       
        List<cashier> info_cashier = new List<cashier>();
        List<employee> info_employee = new List<employee>();
        today info_today = new today();
        List<string> str_list = new List<string>();

        int index_attendence_employee;
        int index_attendence_month;
        Functions boss = new Functions();
        List<attendence> info_attendence = new List<attendence>();
        int index_employee_name;
        private void button2_Click(object sender, EventArgs e)
        {




            if (txt_name_pass.Text == info_today.name_admin && txt_pass.Text == info_today.password_admin)
            {
                panel1.Visible = true;
                pnl_password.Visible = false;
                txt_name_pass.Text = "";
                txt_pass.Text = "";
            }
            else
            {
                MessageBox.Show("wrong password!!!");
                txt_name_pass.Text = "";
                txt_pass.Text = "";
            }
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button61.Enabled = true;
            tottal = 0;
            bill obj_bill = new bill();


            if (txt_num_cut.Text != "" && txt_price_cut.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_cut.Text);
                xd2 = Convert.ToInt32(txt_price_cut.Text);
                tottal += xd1* xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "قص";
                    obj_bill.orders.Add(o);
                }

            }

            if (txt_num_hair.Text != "" && txt_price_hair.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_hair.Text);
                xd2 = Convert.ToInt32(txt_price_hair.Text);
                tottal += xd1* xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "شعر";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_wash_head.Text != "" && txt_price_wash_head.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_wash_head.Text);
                xd2 = Convert.ToInt32(txt_price_wash_head.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "غسيل شعر";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_face_eyebrown.Text != "" && txt_price_face_eyebrown.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_face_eyebrown.Text);
                xd2 = Convert.ToInt32(txt_price_face_eyebrown.Text);
                tottal += xd2*xd1;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "وش وحواجب";
                    obj_bill.orders.Add(o);
                }
            }
           
            if (txt_num_eyebrown.Text != "" && txt_price_eyebrown.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_eyebrown.Text);
                xd2 = Convert.ToInt32(txt_price_eyebrown.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "حواجب";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_color.Text != "" && txt_price_color.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_color.Text);
                xd2 = Convert.ToInt32(txt_price_color.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "صبغه";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_frq_eyebrown.Text != "" && txt_price_frq_eyebrown.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_frq_eyebrown.Text);
                xd2 = Convert.ToInt32(txt_price_frq_eyebrown.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "فرق وش وحواجب";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_hna.Text != "" && txt_price_hna.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_hna.Text);
                xd2 = Convert.ToInt32(txt_price_hna.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "حنه";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_leg_padiquer.Text != "" && txt_price_leg_padiquer.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_leg_padiquer.Text);
                xd2 = Convert.ToInt32(txt_price_leg_padiquer.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "باديكير رجل";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_hand_padiquer.Text != "" && txt_price_hand_padiquer.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_hand_padiquer.Text);
                xd2 = Convert.ToInt32(txt_price_hand_padiquer.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "باديكير ايد";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_makeup.Text != "" && txt_price_makeup.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_makeup.Text);
                xd2 = Convert.ToInt32(txt_price_makeup.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "ميك اب";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_forma.Text != "" && txt_price_forma.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_forma.Text);
                xd2 = Convert.ToInt32(txt_price_forma.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "فورمه";
                    obj_bill.orders.Add(o);
                }
            }
            if (txt_num_brotien.Text != "" && txt_price_brotien.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_brotien.Text);
                xd2 = Convert.ToInt32(txt_price_brotien.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "جلسه تنضيف بشره";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox2.Text);
                xd2 = Convert.ToInt32(textBox1.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "مانيكير";
                    obj_bill.orders.Add(o);
                }

            }
            if (textBox8.Text != "" && textBox7.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox8.Text);
                xd2 = Convert.ToInt32(textBox7.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "جلسه حمام كريم";
                    obj_bill.orders.Add(o);
                }

            }
            if (textBox10.Text != "" && textBox9.Text != "" && textBox11.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox10.Text);
                xd2 = Convert.ToInt32(textBox9.Text);
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = textBox11.Text;
                    obj_bill.orders.Add(o);
                }

            }
            //**************************************
            
            if (textBox20.Text != "" && textBox14.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox20.Text);//#num
                xd2 = Convert.ToInt32(textBox14.Text);//price
                tottal += xd1*xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "وش بدون حواجب";
                    obj_bill.orders.Add(o);
                }
            }

            if (textBox22.Text != "" && textBox18.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox22.Text);//#num
                xd2 = Convert.ToInt32(textBox18.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "رسم حنه";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox21.Text != "" && textBox15.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox21.Text);//#num
                xd2 = Convert.ToInt32(textBox15.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "تركيب اظافر";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox19.Text != "" && textBox6.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox19.Text);//#num
                xd2 = Convert.ToInt32(textBox6.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "قص قوصه";
                    obj_bill.orders.Add(o);
                }
            }
            //****************************************

            if (textBox36.Text != "" && textBox31.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox36.Text);//#num
                xd2 = Convert.ToInt32(textBox31.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "تركيب رموش";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox35.Text != "" && textBox30.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox35.Text);//#num
                xd2 = Convert.ToInt32(textBox30.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "غسيل لاكمي";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox34.Text != "" && textBox29.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox34.Text);//#num
                xd2 = Convert.ToInt32(textBox29.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "شنب";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox38.Text != "" && textBox33.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox38.Text);//#num
                xd2 = Convert.ToInt32(textBox33.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "ماسك";
                    obj_bill.orders.Add(o);
                }
            }


            //****************************************
            if (textBox37.Text != "" && textBox32.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox37.Text);//#num
                xd2 = Convert.ToInt32(textBox32.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "حمام مغربي";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox28.Text != "" && textBox27.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox28.Text);//#num
                xd2 = Convert.ToInt32(textBox27.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "سويت";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox25.Text != "" && textBox23.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox25.Text);//#num
                xd2 = Convert.ToInt32(textBox23.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "فرق شعر";
                    obj_bill.orders.Add(o);
                }
            }
            if (textBox26.Text != "" && textBox24.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox26.Text);//#num
                xd2 = Convert.ToInt32(textBox24.Text);//price
                tottal += xd1 * xd2;
                for (int uiu = 0; uiu < xd1; uiu++)
                {
                    one_order o = new one_order();
                    o.its_price = xd2;
                    o.type_order = "فرق وش";
                    obj_bill.orders.Add(o);
                }
            }



            //**************************************

            if (obj_bill.orders.Count == 0)
            {
                MessageBox.Show("لا يوجد شي للطباعه");

            }
            else if (obj_bill.orders.Count > 51)
            {
                MessageBox.Show("عدد الفواتير يجب ان يقل عن 5 فواتير");
            }
            else
            {
                bill temp = new bill();
                //video youtube insert db

                int temp_index_cashier = index_cashier;
                for (int i = 0; i < obj_bill.orders.Count; i++)
                {
                    temp.orders.Add(obj_bill.orders[i]);
                    if (temp.orders.Count == 10 || (i + 1) == obj_bill.orders.Count)
                    {
                        temp.dat = DateTime.Now;
                        temp.name_cashier = info_cashier[temp_index_cashier].name;
                        info_today.num_bills++;
                        temp.num_this_bill = info_today.num_bills;
                        await Task.Run(() => print_bill_client_savedb(ref temp, ref info_today, temp_index_cashier));
          
                    }
                }
                if (pnl_print.Visible == true)
                {
                    pnl_print.Visible = false;
                    panel2.Visible = true;
                }

            }
            button3.Enabled = true;
            button61.Enabled = true;
        }
        void print_bill_client_without_savedb( bill temp)
        {
            System.IO.File.Copy(Environment.CurrentDirectory + "/att/1.xls", Environment.CurrentDirectory + "/" + info_today.num_bills + ".xls");

           
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Environment.CurrentDirectory + "/" + info_today.num_bills + ".xls");



            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            x.Range["C8"].Value = temp.dat.Month.ToString() + "/" + temp.dat.Day.ToString() + "/" + temp.dat.Year.ToString();
            x.Range["I8"].Value = temp.dat.Month.ToString() + "/" + temp.dat.Day.ToString() + "/" + temp.dat.Year.ToString();

            x.Range["C7"].Value = temp.num_this_bill.ToString();
            x.Range["j9"].Value = temp.name_cashier;
            x.Range["d9"].Value = temp.name_cashier;
            x.Range["k8"].Value = temp.dat.Hour.ToString() + ":" + temp.dat.Minute.ToString();
            x.Range["e8"].Value = temp.dat.Hour.ToString() + ":" + temp.dat.Minute.ToString();

           
            for (int k = 0; k < temp.orders.Count; k++)
            {
                byte[] byt = System.Text.Encoding.UTF8.GetBytes(temp.orders[k].type_order);
                // convert the byte array to a Base64 string
                string strModified = Convert.ToBase64String(byt);
                 x.Range["B" + (k + 11).ToString()].Value = temp.orders[k].type_order;
                x.Range["c" + (k + 11).ToString()].Value = temp.orders[k].its_price.ToString();
                x.Range["D23"].Value = " ";


            }

          
            // Print out 1 copy to the default printer:
            sheet.PrintOut(1, 2, 1, false, Type.Missing, false, false, Type.Missing);


            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();

         
            // PrintMyExcelFile__1(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + info_today.time_today.Year.ToString() + "-" + info_today.time_today.Month.ToString() + "-" + info_today.time_today.Day.ToString() + "/" + info_today.num_bills + ".xls");


            System.IO.File.Delete(Environment.CurrentDirectory + "/" + info_today.num_bills + ".xls");
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }
        void print_bill_client_savedb(ref bill temp, ref today info_today, int temp_index_cashier)
        {
            System.IO.File.Copy(Environment.CurrentDirectory + "/att/1.xls", Environment.CurrentDirectory + "/" + info_today.num_bills + ".xls");

            info_today.bills_of_today.Add(temp);

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Environment.CurrentDirectory + "/" + info_today.num_bills + ".xls");



            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            x.Range["C8"].Value = temp.dat.Month.ToString() + "/" + temp.dat.Day.ToString() + "/" + temp.dat.Year.ToString();
            x.Range["I8"].Value = temp.dat.Month.ToString() + "/" + temp.dat.Day.ToString() + "/" + temp.dat.Year.ToString();

            x.Range["C7"].Value = info_today.num_bills.ToString();
            x.Range["j9"].Value = temp.name_cashier;
            x.Range["d9"].Value = temp.name_cashier;
            x.Range["k8"].Value = temp.dat.Hour.ToString() + ":" + temp.dat.Minute.ToString();
            x.Range["e8"].Value = temp.dat.Hour.ToString() + ":" + temp.dat.Minute.ToString();

            con.Open();
            SQLiteCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;

            for (int k = 0; k < temp.orders.Count; k++)
            {
                byte[] byt = System.Text.Encoding.UTF8.GetBytes(temp.orders[k].type_order);
                // convert the byte array to a Base64 string
                string strModified = Convert.ToBase64String(byt);
                cmd.CommandText += "insert into [mynewtable] (bill_cashier_name,id,type,price,dat) values('" + info_cashier[temp_index_cashier].name + "','" + temp.num_this_bill.ToString() + "','" + strModified + "','" + temp.orders[k].its_price.ToString() + "','" + temp.dat.ToString() + "');";
                x.Range["B" + (k + 11).ToString()].Value = temp.orders[k].type_order;
                x.Range["c" + (k + 11).ToString()].Value = temp.orders[k].its_price.ToString();
                x.Range["D23"].Value = " ";


            }

            cmd.ExecuteNonQuery();
            con.Close();
            // Print out 1 copy to the default printer:
            sheet.PrintOut(1, 2, 1, false, Type.Missing, false, false, Type.Missing);


            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();

            temp = new bill();
            // PrintMyExcelFile__1(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + info_today.time_today.Year.ToString() + "-" + info_today.time_today.Month.ToString() + "-" + info_today.time_today.Day.ToString() + "/" + info_today.num_bills + ".xls");


            System.IO.File.Delete(Environment.CurrentDirectory + "/" + info_today.num_bills + ".xls");
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }
        void PrintMyExcelFile__1(string pa)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Open the Workbook:
            string path = pa;
            Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(path);

            // Get the first worksheet.
            // (Excel uses base 1 indexing, not base 0.)
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            // Print out 1 copy to the default printer:
            wb.PrintOut(1, 2, 1, false, Type.Missing, false, false, Type.Missing);

            // Cleanup:
            //  GC.Collect();
            //    GC.WaitForPendingFinalizers();

            //  System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws);

            wb.Close(false, Type.Missing, Type.Missing);
            //  System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);

            excelApp.Quit();
            // System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(    AppDomain.CurrentDomain.BaseDirectory);
        
            if (File.Exists(@"C:\import\import.txt")) { 
            info_today.num_bills = 0;
            con.Open();
            SQLiteCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT COUNT(*) FROM first_time WHERE id IS NOT NULL";
            int rresult = int.Parse(cmd.ExecuteScalar().ToString());
            con.Close();
            if (rresult == 0)
            {
                con.Open();
                SQLiteCommand cmmd = con.CreateCommand();
                cmmd.CommandType = CommandType.Text;
                cmmd.CommandText = "insert into mynewtable_admin (Name_admin,Pass_admin) values('admin','admin')";

                // cmmd.CommandText = "alter table [mynewtable_admin] drop column Name_admin;";
                cmmd.ExecuteNonQuery();

                con.Close();
                con.Open();
                cmmd = con.CreateCommand();
                cmmd.CommandType = CommandType.Text;
                cmmd.CommandText = "insert into first_time (id) values(11)";

                // cmmd.CommandText = "alter table [mynewtable_admin] drop column Name_admin;";
                cmmd.ExecuteNonQuery();

                con.Close();
            }

            //--------------------------------

            con.Open();
            cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT COUNT(*) FROM [mynewtable] WHERE id IS NOT NULL";
            int result = int.Parse(cmd.ExecuteScalar().ToString());
            con.Close();
            con.Open();
            //----------
            cmd = new SQLiteCommand("select pass_admin from [mynewtable_admin]", con);
            SQLiteDataReader dr0 = cmd.ExecuteReader();
            int h = 0;
            while (dr0.Read())
            {
                h++;
                info_today.password_admin = dr0.GetString(dr0.GetOrdinal("pass_admin"));
            }

            con.Close();
            con.Open();
            cmd = new SQLiteCommand("select name_admin from [mynewtable_admin]", con);
            con.Close();
            con.Open();
            SQLiteDataReader dr7 = cmd.ExecuteReader();

            dr7.Read();

            info_today.name_admin = dr7.GetString(dr7.GetOrdinal("name_admin"));

            //-------4
            con.Close();

            con.Open();
            cmd = new SQLiteCommand("SELECT COUNT(*) FROM [mynewtable_cashier] WHERE cashier_name IS NOT NULL", con);
            int result_cashie = int.Parse(cmd.ExecuteScalar().ToString());
            List<string> name_cashier = new List<string>();
            con.Close();
            if (result_cashie != 0)
            {

                con.Open();
                cmd = new SQLiteCommand("select cashier_name from [mynewtable_cashier]", con);
                SQLiteDataReader dr8 = cmd.ExecuteReader();

                while (dr8.Read())
                {

                    name_cashier.Add(dr8.GetString(dr8.GetOrdinal("cashier_name")));
                }
                con.Close();
                con.Open();
                cmd = new SQLiteCommand("select cashier_pass from [mynewtable_cashier]", con);
                SQLiteDataReader dr9 = cmd.ExecuteReader();
                int oop = 0;
                while (dr9.Read())
                {
                    cashier temp = new cashier();

                    temp.password = dr9.GetString(dr9.GetOrdinal("cashier_pass"));
                    temp.name = name_cashier[oop];
                    info_cashier.Add(temp);
                    oop++;
                }
                con.Close();

            }
            con.Open();
            cmd = new SQLiteCommand("SELECT COUNT(*) FROM [mynewtable_time] WHERE time IS NOT NULL", con);
            int re = int.Parse(cmd.ExecuteScalar().ToString());
            con.Close();
            if (re == 0)
            {
                info_today.time_today = DateTime.Now;
                con.Open();
                SQLiteCommand cmmd = con.CreateCommand();
                cmmd.CommandType = CommandType.Text;
                cmmd.CommandText = "insert into mynewtable_time (time) values('" + info_today.time_today.ToString() + "')";

                // cmmd.CommandText = "alter table [mynewtable_admin] drop column Name_admin;";
                cmmd.ExecuteNonQuery();

                con.Close();
            }
            else
            {
                con.Open();
                cmd = new SQLiteCommand("select time from [mynewtable_time]", con);
                SQLiteDataReader drr = cmd.ExecuteReader();

                drr.Read();
                    string r = drr.GetString(drr.GetOrdinal("time"));
                    if (r[r.Length - 1] == 'ص')
                    {

                        r = r.Remove(r.Length - 1);
                        r += "AM";

                    }
                    else if (r[r.Length - 1] == 'م')
                    {
                        r = r.Remove(r.Length - 1);
                        r += "PM";

                    }
                    DateTime tyty = new DateTime();
                    if (DateTime.TryParseExact(r, "dd/MM/yyyy hh:mm:ss tt", CultureInfo.GetCultureInfo("en-us"), DateTimeStyles.None, out tyty))
                    {
                        //       "03/12/2019 10:42:20 ص"

                    }
                    else if (DateTime.TryParseExact(r, "dd/MM/yyyy hh:mm:ss tt", CultureInfo.GetCultureInfo("en-uk"), DateTimeStyles.None, out tyty))
                    {

                    }
                    else
                    {

                        tyty = Convert.ToDateTime(r);

                    }
               

                    info_today.time_today = tyty;
                    

                con.Close();


            }



            if (result == 0)// if result equals zero, then the table is empty
            {



                info_today.num_bills = 0;
                info_today.num_employees = 0;


            }
            else
            {
                List<string> id = new List<string>();
                List<string> type = new List<string>();
                List<string> price = new List<string>();
                List<string> name = new List<string>();
                List<DateTime> tim = new List<DateTime>();
                con.Open();
                cmd = new SQLiteCommand("select dat from [mynewtable]", con);
                SQLiteDataReader drr1 = cmd.ExecuteReader();

                while (drr1.Read())
                {
                    string str = drr1.GetString(drr1.GetOrdinal("dat"));
                       
                        if (str[str.Length-1] == 'ص')
                        {

                         str=   str.Remove(str.Length - 1);
                            str += "AM";
                         
                        }
                        else if (str[str.Length - 1] == 'م')
                        {
                          str=  str.Remove(str.Length - 1);
                            str += "PM";
                          
                        }
                        DateTime tyty = new DateTime();
                        if(DateTime.TryParseExact(str, "dd/MM/yyyy hh:mm:ss tt", CultureInfo.GetCultureInfo("en-us"),DateTimeStyles.None,out tyty)){
                            //       "03/12/2019 10:42:20 ص"
                         
                        }
                       else if (DateTime.TryParseExact(str, "dd/MM/yyyy hh:mm:ss tt", CultureInfo.GetCultureInfo("en-uk"), DateTimeStyles.None, out tyty))
                        {
                           
                        }
                        else
                        {
                           
                            tyty = Convert.ToDateTime(str);
                           
                        }
                    tim.Add(tyty);

                        //   DateTime.TryParseExact(str, "dd/MM/yy hh:mm:ss", CultureInfo.GetCultureInfo("ar-sa"

                    }
                    con.Close();
                con.Open();
                cmd = new SQLiteCommand("select id from [mynewtable]", con);
                SQLiteDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {

                    id.Add(dr.GetString(dr.GetOrdinal("id")));



                }
                con.Close();
                con.Open();
                cmd = new SQLiteCommand("select type from [mynewtable]", con);
                SQLiteDataReader dr1 = cmd.ExecuteReader();
                while (dr1.Read())
                {
                    string strModified = dr1.GetString(dr1.GetOrdinal("type"));

                    byte[] b = Convert.FromBase64String(strModified);
                    string strOriginal = System.Text.Encoding.UTF8.GetString(b);
                    type.Add(strOriginal);
                }
                con.Close();
                con.Open();
                cmd = new SQLiteCommand("select bill_cashier_name from [mynewtable]", con);
                SQLiteDataReader dr2 = cmd.ExecuteReader();
                while (dr2.Read())
                {
                    name.Add(dr2.GetString(dr2.GetOrdinal("bill_cashier_name")));

                }
                con.Close();
                con.Open();
                cmd = new SQLiteCommand("select price from [mynewtable]", con);
                SQLiteDataReader dr3 = cmd.ExecuteReader();
                int i = 1;
                int prev_id = Convert.ToInt32(id[0]);
                int j = 0;
                bill tempo = new bill();

                while (dr3.Read())
                {
                    price.Add(dr3.GetString(dr3.GetOrdinal("price")));

                    one_order tempo_order = new one_order();
                    tempo_order.id = Convert.ToInt32(id[j]);
                    tempo_order.its_price = Convert.ToInt32(price[j]);
                    tempo_order.type_order = type[j];
                    j++;

                    tempo.orders.Add(tempo_order);
                    if (j != id.Count)
                    {
                        i = Convert.ToInt32(id[j]);
                        if (prev_id != i)
                        {
                            tempo.num_this_bill = prev_id;
                            tempo.name_cashier = name[j - 1];
                            tempo.dat = tim[j - 1];

                            //   tempo.orders.Add(tempo_order);
                            prev_id = i;
                            info_today.bills_of_today.Add(tempo);
                            info_today.num_bills=tempo.num_this_bill;
                            tempo = new bill();
                        }
                    }
                    else
                    {//دي اخر سطر
                        tempo.num_this_bill = prev_id;
                        tempo.name_cashier = name[j - 1];
                        tempo.dat = tim[j - 1];
                        //  tempo.orders.Add(tempo_order);
                        prev_id = i;
                        info_today.bills_of_today.Add(tempo);
                        info_today.num_bills=tempo.num_this_bill;
                        tempo = new bill();
                    }

                }

                con.Close();



            }
            //--تعديل
            //  info_today.num_bills = 70;

            button1.Visible = false;
            pnl_menu.Visible = true;
        }
    }

        private void button4_Click(object sender, EventArgs e)
        {
           
               pnl_menu.Visible = false;
            button1.Visible = false;
            pnl_password.Visible = true;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            panel1.Visible = false;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox10.Text = "";
            textBox11.Text = "";
            textBox9.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            txt_num_brotien.Text = "";
            txt_num_color.Text = "";
            txt_num_cut.Text = "";
            txt_num_eyebrown.Text = "";
          
            txt_num_face_eyebrown.Text = "";
            txt_num_forma.Text = "";
            txt_num_frq_eyebrown.Text = "";
            txt_num_hair.Text = "";
            txt_num_hand_padiquer.Text = "";
            txt_num_hna.Text = "";
            txt_num_leg_padiquer.Text = "";
            txt_num_makeup.Text = "";
            txt_num_wash_head.Text = "";
            txt_price_brotien.Text = "";
            txt_price_color.Text = "";
            txt_price_cut.Text = "";
            txt_price_eyebrown.Text = "";
            
            txt_price_face_eyebrown.Text = "";
            txt_price_forma.Text = "";
            txt_price_frq_eyebrown.Text = "";
            txt_price_hair.Text = "";
            txt_price_hand_padiquer.Text = "";
            txt_price_hna.Text = "";
            txt_price_leg_padiquer.Text = "";
            txt_price_makeup.Text = "";
            txt_price_wash_head.Text = "";
            textBox20.Text = "";
            textBox14.Text = "";
            textBox22.Text = "";
            textBox18.Text = "";
            textBox21.Text = "";
            textBox15.Text = "";
            textBox19.Text = "";
            textBox6.Text = "";
            textBox36.Text = "";
            textBox31.Text = "";
            textBox35.Text = "";
            textBox30.Text = "";
            textBox34.Text = "";
            textBox29.Text = "";
            textBox38.Text = "";
            textBox33.Text = "";
            textBox37.Text = "";
            textBox32.Text = "";
            textBox28.Text = "";
            textBox27.Text = "";
            textBox25.Text = "";
            textBox23.Text = "";
            textBox26.Text = "";
            textBox24.Text = "";
            textBox39.Text = "";
            pnl_print.Visible = true;
            panel2.Visible = false;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            pnl_password.Visible = false;
            pnl_menu.Visible = true;
            txt_name_pass.Text = "";
            txt_pass.Text = "";
        }

        private void button13_Click(object sender, EventArgs e)
        {
            index_cashier = -1;
            panel2.Visible = false;
            pnl_menu.Visible = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            pnl_menu.Visible = true;
            textBox3.Text = "";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
          
        }

        private void button22_Click(object sender, EventArgs e)
        {
            if (textBox17.Text == "" || textBox16.Text == "")
            {
                MessageBox.Show("خطا يجب ان تجرب مره اخري!!!");
                textBox16.Text = "";
                textBox17.Text = "";
            }
            else
            {
                int yaya = -1;
                for (int i = 0; i < info_cashier.Count; i++)
                {
                    if (info_cashier[i].name == textBox17.Text)
                    {
                        yaya = i;
                    }

                }
                if (yaya != -1)
                {
                    MessageBox.Show("خطا  يجب عليك تغيير  الاسم لانه مستخدم من قبل");
                }
                else
                {
                    cashier temp = new cashier();
                    temp.name = textBox17.Text;
                    temp.password = textBox16.Text;
                    info_cashier.Add(temp);
                    con.Open();
                    SQLiteCommand cmmd = con.CreateCommand();
                    cmmd.CommandType = CommandType.Text;
                    cmmd.CommandText = "insert into mynewtable_cashier (cashier_name,cashier_pass) values('" + temp.name + "','" + temp.password + "')";

                    // cmmd.CommandText = "alter table [mynewtable_admin] drop column Name_admin;";
                    cmmd.ExecuteNonQuery();

                    con.Close();

                    panel1.Visible = true;
                    panel3.Visible = false;
                    textBox16.Text = "";
                    textBox17.Text = "";
                    MessageBox.Show("تم اضافه الكاشير بنجاح!!!");
                }
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel3.Visible = false;
            textBox16.Text = "";
            textBox17.Text = "";
        }

        /* private void button24_Click(object sender, EventArgs e)
          {
              if (textBox19.Text == "" || textBox18.Text == "" || textBox21.Text == "")
              {
                  textBox19.Text = "";
                  textBox18.Text = "";

                  textBox21.Text = "";
                  MessageBox.Show("خطا يجب ان تكتب كل البيانات كامله!!!");

              }
              else
              {
                  int yaya = -1;
                  for (int i = 0; i < info_employee.Count; i++)
                  {
                      if (info_employee[i].name == textBox19.Text)
                      {
                          yaya = i;
                      }

                  }
                  if (yaya != -1)
                  {
                      MessageBox.Show(" خطا  يجب عليك تغيير  الاسم لانه مستخدم من قبل مستخدم اخر");
                  }
                  else
                  {
                      employee temp = new employee();
                      temp.name = textBox19.Text;
                      temp.id = textBox18.Text;
                      temp.address = textBox21.Text;

                      temp.time_start_work = DateTime.Now;

                      info_employee.Add(temp);
                      onebyone oo = new onebyone();
                      oo.name = temp.name;
                      info_attendence[info_attendence.Count - 1].employees_month.Add(oo);

                      textBox19.Text = "";
                      textBox18.Text = "";

                      textBox21.Text = "";
                      MessageBox.Show("تم اضافه الموظف بنجاح!!!");
                      panel1.Visible = true;
                      panel4.Visible = false;
                  }
              }
          }
          */
        private void button6_Click(object sender, EventArgs e)
        {
            textBox12.Text = "";

            textBox13.Text = "";

            pnl_menu.Visible = false;
            panel22.Visible = true;

        }




        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            panel10.Visible = true; panel1.Visible = false;
        }

        private void txt_num_face_TextChanged(object sender, EventArgs e)
        {

        }



        private void button15_Click(object sender, EventArgs e)
        {
            panel5.Visible = true;
            panel1.Visible = false;
        }

        private void button25_Click(object sender, EventArgs e)
        {
            panel5.Visible = false;
            panel1.Visible = true;
        }





        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }


        private void panel9_Paint(object sender, PaintEventArgs e)
        {

        }



        private void button36_Click(object sender, EventArgs e)
        {
            con.Open();
            SQLiteCommand cmmd = con.CreateCommand();
            cmmd.CommandType = CommandType.Text;
            cmmd.CommandText = "UPDATE mynewtable_admin   SET Name_admin = '" + textBox5.Text + "', Pass_admin = '" + textBox4.Text + "' Where Name_admin = '" + info_today.name_admin + "' and Pass_admin = '" + info_today.password_admin + "'";
            //"insert into mynewtable_admin (Name_admin,Pass_admin) values('admin','admin')";



            cmmd.ExecuteNonQuery();

            con.Close();
            info_today.name_admin = textBox5.Text;
            info_today.password_admin = textBox4.Text;
            textBox5.Text = "";
            textBox4.Text = "";
            panel1.Visible = true; panel10.Visible = false;
        }

        private void button35_Click(object sender, EventArgs e)
        {
            panel1.Visible = true; panel10.Visible = false;
            textBox5.Text = "";
            textBox4.Text = "";

        }




        private void button33_Click(object sender, EventArgs e)
        {
            select_index = -1;
            int hy = -1;
            listView7.Items.Clear();
            for (int i = 0; i < info_cashier.Count; i++)
            {
                ListViewItem v = new ListViewItem(info_cashier[i].name);
                v.SubItems.Add(info_cashier[i].password);
                listView7.Items.Add(v);
                hy = 0;
            }
            if (hy == -1)
            {
                MessageBox.Show("لاتوجد بيانات لعرضها !!");
            }
            else
            {
                panel13.Visible = true;
                panel5.Visible = false;
            }
        }
        private static string[] GetFileNames(string path, string filter)
        {
            string[] files = Directory.GetFiles(path, filter);
            for (int i = 0; i < files.Length; i++)
                files[i] = Path.GetFileName(files[i]);
            return files;
        }




        int yhf = -1;







        private void button41_Click(object sender, EventArgs e)
        {
            panel13.Visible = false;
            panel1.Visible = true;
        }



        private void button50_Click(object sender, EventArgs e)
        {
            select_index = -1;
            listView10.Items.Clear();
            for (int i = 0; i < info_cashier.Count; i++)
            {
                ListViewItem v = new ListViewItem(info_cashier[i].name);
                v.SubItems.Add(info_cashier[i].password);
                listView10.Items.Add(v);

            }


            panel1.Visible = false;
            panel18.Visible = true;
        }

        private void button53_Click(object sender, EventArgs e)
        {
            if (select_index == 1)
            {
                int u = -1;
                for (int i = 0; i < info_cashier.Count; i++)
                {
                    if (info_cashier[i].name == listView10.SelectedItems[0].SubItems[0].Text && info_cashier[i].password == listView10.SelectedItems[0].SubItems[1].Text)
                    {
                        u = i;
                    }
                }
                if (u == -1)
                {
                    MessageBox.Show("الاسم او الباسورد خطا اعد المحاوله");
                }
                else
                {
                    con.Open();
                    SQLiteCommand cmmd = con.CreateCommand();
                    cmmd.CommandType = CommandType.Text;
                    cmmd.CommandText = "DELETE FROM mynewtable_cashier WHERE cashier_name = '" + info_cashier[u].name + "'";

                    // cmmd.CommandText = "alter table [mynewtable_admin] drop column Name_admin;";
                    cmmd.ExecuteNonQuery();

                    con.Close();
                    info_cashier.Remove(info_cashier[u]);
                    MessageBox.Show("لقد تم المسح بنجاح");
                    panel18.Visible = false;
                    panel1.Visible = true;

                }
            }
            else
            {
                MessageBox.Show("اختار كاشير عشان تقدر تمسحه");
            }
            select_index = -1;
        }

        private void button52_Click(object sender, EventArgs e)
        {
            panel18.Visible = false;
            panel1.Visible = true;
        }










        private void panel22_Paint(object sender, PaintEventArgs e)
        {
          
        }

        private void button60_Click(object sender, EventArgs e)
        {
            index_cashier = -1;
            for (int i = 0; i < info_cashier.Count; i++)
            {
                if (textBox13.Text == info_cashier[i].name && textBox12.Text == info_cashier[i].password)
                {
                    index_cashier = i;
                }
            }
            if (index_cashier == -1)
            {
                textBox12.Text = "";

                textBox13.Text = "";

                MessageBox.Show("خطا حاول مره اخري");
            }
            else
            {
                MessageBox.Show("تم تسجيل الدخول بنجاح");
                panel22.Visible = false;
                panel2.Visible = true;
            }
        }

        private void button59_Click(object sender, EventArgs e)
        {
            textBox12.Text = "";

            textBox13.Text = "";

            pnl_menu.Visible = true;
            panel22.Visible = false;
        }

        private void button61_Click(object sender, EventArgs e)
        {
            pnl_print.Visible = false;
            panel2.Visible = true;
        }
        void fininsh_today()
        {
         string tot   = "";
            int sum = 0, counter = 1;

            int sum_10_bill = 0;
            for (int i = 0; i < info_today.bills_of_today.Count; i++)
            {
                int sum_bill = 0;
                for (int j = 0; j < info_today.bills_of_today[i].orders.Count; j++)
                {
                    sum_bill += info_today.bills_of_today[i].orders[j].its_price;
                    sum += info_today.bills_of_today[i].orders[j].its_price;
                }
                sum_10_bill += sum_bill;
                if (i == info_today.bills_of_today.Count - 1)
                {
                    tot += counter + " = " + sum_10_bill + ".\r\n";
                    sum_10_bill = 0;
                }
                else if (i+1 == counter&&counter%10==0)
                {

                    tot += (counter - 9) + " .. " + counter + " = " + sum_10_bill + ".\r\n";
                    sum_10_bill = 0;
                }
                counter++;
            }
            tot += "Toatal = " + sum + " .";
           
            int iop = 0;
            iop = info_attendence.Count;

            using (ZipFile zip = new ZipFile())
            {
                StreamWriter w = new StreamWriter("total.txt");
                w.Write(tot);
                w.Close();
                ZipEntry e1 = new ZipEntry();


                e1 = zip.AddFile("total.txt");
                e1.Password = "amgad" + info_today.time_today.Year.ToString() + info_today.time_today.Month.ToString() + info_today.time_today.Day.ToString();
                e1.Encryption = EncryptionAlgorithm.WinZipAes256;

                for (int i = 0; i < info_today.bills_of_today.Count; i++)
                {
                    System.IO.File.Copy(Environment.CurrentDirectory + "/att/1.xls", Environment.CurrentDirectory + "/" + info_today.bills_of_today[i].num_this_bill + ".xls");


                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Environment.CurrentDirectory + "/" + info_today.bills_of_today[i].num_this_bill + ".xls");



                    Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    x.Range["C8"].Value = info_today.time_today.Month.ToString() + "/" + info_today.time_today.Day.ToString() + "/" + info_today.time_today.Year.ToString();
                    x.Range["I8"].Value = info_today.time_today.Month.ToString() + "/" + info_today.time_today.Day.ToString() + "/" + info_today.time_today.Year.ToString();

                    x.Range["C7"].Value = info_today.bills_of_today[i].num_this_bill.ToString();
                    x.Range["j9"].Value = info_today.bills_of_today[i].name_cashier.ToString();
                    x.Range["d9"].Value = info_today.bills_of_today[i].name_cashier.ToString();
                    x.Range["k8"].Value = info_today.bills_of_today[i].dat.Hour.ToString() + ":" + info_today.bills_of_today[i].dat.Minute.ToString();
                    x.Range["e8"].Value = info_today.bills_of_today[i].dat.Hour.ToString() + ":" + info_today.bills_of_today[i].dat.Minute.ToString();

                    for (int k = 0; k < info_today.bills_of_today[i].orders.Count; k++)
                    {
                        x.Range["B" + (k + 11).ToString()].Value = info_today.bills_of_today[i].orders[k].type_order;
                        x.Range["c" + (k + 11).ToString()].Value = info_today.bills_of_today[i].orders[k].its_price.ToString();
                        x.Range["D23"].Value = " ";


                    }


                    sheet.Close(true, Type.Missing, Type.Missing);
                    excel.Quit();
                     e1 = new ZipEntry();


                    e1 = zip.AddFile(info_today.bills_of_today[i].num_this_bill + ".xls");
                    e1.Password = "amgad"+ info_today.time_today.Year.ToString() + info_today.time_today.Month.ToString() +info_today.time_today.Day.ToString();
                    e1.Encryption = EncryptionAlgorithm.WinZipAes256;

                }

                if (!Directory.Exists("files"))
                {
                    Directory.CreateDirectory("files");
                }


                zip.Save("files/"+info_today.time_today.Year.ToString() + "-" + info_today.time_today.Month.ToString() + "-" + info_today.time_today.Day.ToString() + ".zip");



            }
            System.IO.File.Delete(Environment.CurrentDirectory + "/" + "total.txt");

            for (int i = 0; i < info_today.bills_of_today.Count; i++)
            {
                System.IO.File.Delete(Environment.CurrentDirectory + "/" + info_today.bills_of_today[i].num_this_bill + ".xls");
            }
            // alter table t drop column col;
            //alter table t set unused(col);
            //**********************
            con.Open();
            // SQLiteCommand cmmd = new SQLiteCommand("truncate table mynewtable", con);
            SQLiteCommand cmmd = con.CreateCommand();
            cmmd.CommandType = CommandType.Text;

            cmmd.CommandText = ("delete from mynewtable ;");
            cmmd.ExecuteNonQuery();
            con.Close();
            con.Open();
            // cmmd = new SQLiteCommand("truncate table mynewtable_time", con);
            cmmd = new SQLiteCommand("delete from mynewtable_time ;", con);
            cmmd.ExecuteNonQuery();
            con.Close();
        }
        protected async void closefun(FormClosingEventArgs e)
        {


            /*   con.Open();
                SqlCommand cmmd = con.CreateCommand();
                cmmd.CommandType = CommandType.Text;
                  cmmd.CommandText = "insert into mynewtable_admin (Name_admin,Pass_admin) values('admin','admin')";

               // cmmd.CommandText = "alter table [mynewtable_admin] drop column Name_admin;";
                cmmd.ExecuteNonQuery();

                con.Close();*/


            base.OnFormClosing(e);
            if (false)
            {
                /* boss.extract_files("amgaDFGHJJlarozmichael154789653");
                 boss.WriteInFileFromList_info_today(info_today, info_employee);
                 boss.WriteInFileFromList_cashier(info_cashier, ref info_today);
                 boss.WriteInFileFromList_employee(info_employee);

                 for (int i = 0; i < info_attendence.Count; i++)
                 {
                     boss.WriteInFileFromList_attendence(info_attendence, str_list[i], i);

                 }
                 FileStream fs = new FileStream(Environment.CurrentDirectory + "/att/info.txt", FileMode.Create);
                 StreamWriter sw = new StreamWriter(fs);
                 for (int i = 0; i < str_list.Count; i++)
                 {
                     sw.WriteLine(str_list[i]);
                 }
                 sw.Close();


                 boss.compress_files("amgaDFGHJJlarozmichael154789653");
                 */
            }
            else
            {
                if (exitt == 1)
                {
                    // Confirm user wants to close
                    switch (MessageBox.Show(this, "هل تود انهاء اليوم ؟؟", "Closing", MessageBoxButtons.YesNoCancel))
                    {
                        case DialogResult.Cancel:
                          
                            exitt = 0;
                            break;

                        case DialogResult.Yes:

                            button19.Enabled = false;
                            button50.Enabled = false;
                            button12.Enabled = false;
                            button15.Enabled = false;
                            button5.Enabled = false;
                            button8.Enabled = false;
                            button10.Enabled = false;
                            button9.Enabled = false;
                            button14.Enabled = false;
                            await Task.Run(() => fininsh_today());

                            this.Close();
                          

                            break;

                        default:
                           
                            exitt = 0;
                          
                            break;


                    }
                }
                else
                {

                }
            }

        }

        private void listView9_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView8_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void panel14_Paint(object sender, PaintEventArgs e)
        {

        }

        private void listView10_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView11_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView12_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView5_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView2_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView15_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;
        }

        private void listView16_MouseClick(object sender, MouseEventArgs e)
        {
            select_index = 1;


        }
        FormClosingEventArgs ee;
        private void button14_Click(object sender, EventArgs e)
        {
            exitt = 1;

            closefun(ee);


        }

        private void button23_Click(object sender, EventArgs e)
        {

        }
     void   save_files()
        {
            string tot = "";
            int sum = 0, counter = 1;

            int sum_10_bill = 0;
            for (int i = 0; i < info_today.bills_of_today.Count; i++)
            {
                int sum_bill = 0;
                for (int j = 0; j < info_today.bills_of_today[i].orders.Count; j++)
                {
                    sum_bill += info_today.bills_of_today[i].orders[j].its_price;
                    sum += info_today.bills_of_today[i].orders[j].its_price;
                }
                sum_10_bill += sum_bill;
                if (i == info_today.bills_of_today.Count - 1)
                {
                    tot += counter + " = " + sum_10_bill + ".\r\n";
                    sum_10_bill = 0;
                }
                else if (i +1== counter && counter % 10 == 0)
                {

                    tot += (counter - 9) + " .. " + counter + " = " + sum_10_bill + ".\r\n";
                    sum_10_bill = 0;
                }
                counter++;
            }
            tot += "Toatal = " + sum + " .";
            using (ZipFile zip = new ZipFile())
            {
                StreamWriter w = new StreamWriter("total.txt");
                w.Write(tot);
                w.Close();
                ZipEntry e1 = new ZipEntry();


                e1 = zip.AddFile("total.txt");
                e1.Password = "amgad" + info_today.time_today.Year.ToString() + info_today.time_today.Month.ToString() + info_today.time_today.Day.ToString();
                e1.Encryption = EncryptionAlgorithm.WinZipAes256;
                for (int i = 0; i < info_today.bills_of_today.Count; i++)
                {
                    System.IO.File.Copy(Environment.CurrentDirectory + "/att/1.xls", Environment.CurrentDirectory + "/" + info_today.bills_of_today[i].num_this_bill + ".xls");


                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(Environment.CurrentDirectory + "/" + info_today.bills_of_today[i].num_this_bill + ".xls");



                    Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    x.Range["C8"].Value = info_today.time_today.Month.ToString() + "/" + info_today.time_today.Day.ToString() + "/" + info_today.time_today.Year.ToString();
                    x.Range["I8"].Value = info_today.time_today.Month.ToString() + "/" + info_today.time_today.Day.ToString() + "/" + info_today.time_today.Year.ToString();

                    x.Range["C7"].Value = info_today.bills_of_today[i].num_this_bill.ToString();
                    x.Range["j9"].Value = info_today.bills_of_today[i].name_cashier.ToString();
                    x.Range["d9"].Value = info_today.bills_of_today[i].name_cashier.ToString();
                    x.Range["k8"].Value = info_today.bills_of_today[i].dat.Hour.ToString() + ":" + info_today.bills_of_today[i].dat.Minute.ToString();
                    x.Range["e8"].Value = info_today.bills_of_today[i].dat.Hour.ToString() + ":" + info_today.bills_of_today[i].dat.Minute.ToString();

                    for (int k = 0; k < info_today.bills_of_today[i].orders.Count; k++)
                    {
                        x.Range["B" + (k + 11).ToString()].Value = info_today.bills_of_today[i].orders[k].type_order;
                        x.Range["c" + (k + 11).ToString()].Value = info_today.bills_of_today[i].orders[k].its_price.ToString();
                        x.Range["D23"].Value = " ";


                    }


                    sheet.Close(true, Type.Missing, Type.Missing);
                    excel.Quit();
                     e1 = new ZipEntry();


                    e1 = zip.AddFile(info_today.bills_of_today[i].num_this_bill + ".xls");
                    e1.Password = "amgad" + info_today.time_today.Year.ToString() + info_today.time_today.Month.ToString() + info_today.time_today.Day.ToString();
                    e1.Encryption = EncryptionAlgorithm.WinZipAes256;

                }




                zip.Save(info_today.time_today.Year.ToString() + "-" + info_today.time_today.Month.ToString() + "-" + info_today.time_today.Day.ToString() + ".zip");



            }
            System.IO.File.Delete(Environment.CurrentDirectory + "/" + "total.txt");

            for (int i = 0; i < info_today.bills_of_today.Count; i++)
            {
                System.IO.File.Delete(Environment.CurrentDirectory + "/" + info_today.bills_of_today[i].num_this_bill + ".xls");
            }
        }
        private async void button8_Click(object sender, EventArgs e)
        {
            await Task.Run(() => save_files());
            MessageBox.Show("تم اانتهاء من الطباعه");
        }

        private void button9_Click(object sender, EventArgs e)
        {
           
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            textBox3.Text = info_today.num_bills.ToString();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {if(info_today.num_bills<= Convert.ToInt32(textBox3.Text))
                info_today.num_bills = Convert.ToInt32(textBox3.Text);
                else MessageBox.Show("اختار رقم اكبر من اخر فاتوره");
            }
        }

        private void pnl_print_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            tottal = 0;
            if (txt_num_cut.Text != "" && txt_price_cut.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_cut.Text);
                xd2 = Convert.ToInt32(txt_price_cut.Text);
                tottal += xd1 * xd2;
               

            }

            if (txt_num_hair.Text != "" && txt_price_hair.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_hair.Text);
                xd2 = Convert.ToInt32(txt_price_hair.Text);
                tottal += xd1 * xd2;
               
            }
            if (txt_num_wash_head.Text != "" && txt_price_wash_head.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_wash_head.Text);
                xd2 = Convert.ToInt32(txt_price_wash_head.Text);
                tottal += xd1 * xd2;
               
            }
            if (txt_num_face_eyebrown.Text != "" && txt_price_face_eyebrown.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_face_eyebrown.Text);
                xd2 = Convert.ToInt32(txt_price_face_eyebrown.Text);
                tottal += xd2 * xd1;
              
            }
       
            if (txt_num_eyebrown.Text != "" && txt_price_eyebrown.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_eyebrown.Text);
                xd2 = Convert.ToInt32(txt_price_eyebrown.Text);
                tottal += xd1 * xd2;
               
            }
            if (txt_num_color.Text != "" && txt_price_color.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_color.Text);
                xd2 = Convert.ToInt32(txt_price_color.Text);
                tottal += xd1 * xd2;
               
            }
            if (txt_num_frq_eyebrown.Text != "" && txt_price_frq_eyebrown.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_frq_eyebrown.Text);
                xd2 = Convert.ToInt32(txt_price_frq_eyebrown.Text);
                tottal += xd1 * xd2;
              
            }
            if (txt_num_hna.Text != "" && txt_price_hna.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_hna.Text);
                xd2 = Convert.ToInt32(txt_price_hna.Text);
                tottal += xd1 * xd2;
              
            }
            if (txt_num_leg_padiquer.Text != "" && txt_price_leg_padiquer.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_leg_padiquer.Text);
                xd2 = Convert.ToInt32(txt_price_leg_padiquer.Text);
                tottal += xd1 * xd2;
              
            }
            if (txt_num_hand_padiquer.Text != "" && txt_price_hand_padiquer.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_hand_padiquer.Text);
                xd2 = Convert.ToInt32(txt_price_hand_padiquer.Text);
                tottal += xd1 * xd2;
              
            }
            if (txt_num_makeup.Text != "" && txt_price_makeup.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_makeup.Text);
                xd2 = Convert.ToInt32(txt_price_makeup.Text);
                tottal += xd1 * xd2;
           
            }
            if (txt_num_forma.Text != "" && txt_price_forma.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_forma.Text);
                xd2 = Convert.ToInt32(txt_price_forma.Text);
                tottal += xd1 * xd2;
              
            }
            if (txt_num_brotien.Text != "" && txt_price_brotien.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(txt_num_brotien.Text);
                xd2 = Convert.ToInt32(txt_price_brotien.Text);
                tottal += xd1 * xd2;
              
            }
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox2.Text);
                xd2 = Convert.ToInt32(textBox1.Text);
                tottal += xd1 * xd2;
            

            }
            if (textBox8.Text != "" && textBox7.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox8.Text);
                xd2 = Convert.ToInt32(textBox7.Text);
                tottal += xd1 * xd2;
            

            }
            if (textBox10.Text != "" && textBox9.Text != "" && textBox11.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox10.Text);
                xd2 = Convert.ToInt32(textBox9.Text);
                tottal += xd1 * xd2;
           

            }
            //**************************************

            if (textBox20.Text != "" && textBox14.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox20.Text);//#num
                xd2 = Convert.ToInt32(textBox14.Text);//price
                tottal += xd1 * xd2;
              
            }

            if (textBox22.Text != "" && textBox18.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox22.Text);//#num
                xd2 = Convert.ToInt32(textBox18.Text);//price
                tottal += xd1 * xd2;
             
            }
            if (textBox21.Text != "" && textBox15.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox21.Text);//#num
                xd2 = Convert.ToInt32(textBox15.Text);//price
                tottal += xd1 * xd2;
               
            }
            if (textBox19.Text != "" && textBox6.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox19.Text);//#num
                xd2 = Convert.ToInt32(textBox6.Text);//price
                tottal += xd1 * xd2;
              
            }
            //****************************************

            if (textBox36.Text != "" && textBox31.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox36.Text);//#num
                xd2 = Convert.ToInt32(textBox31.Text);//price
                tottal += xd1 * xd2;
            
            }
            if (textBox35.Text != "" && textBox30.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox35.Text);//#num
                xd2 = Convert.ToInt32(textBox30.Text);//price
                tottal += xd1 * xd2;
           
            }
            if (textBox34.Text != "" && textBox29.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox34.Text);//#num
                xd2 = Convert.ToInt32(textBox29.Text);//price
                tottal += xd1 * xd2;
               
            }
            if (textBox38.Text != "" && textBox33.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox38.Text);//#num
                xd2 = Convert.ToInt32(textBox33.Text);//price
                tottal += xd1 * xd2;
              
            }


            //****************************************
            if (textBox37.Text != "" && textBox32.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox37.Text);//#num
                xd2 = Convert.ToInt32(textBox32.Text);//price
                tottal += xd1 * xd2;
              
            }
            if (textBox28.Text != "" && textBox27.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox28.Text);//#num
                xd2 = Convert.ToInt32(textBox27.Text);//price
                tottal += xd1 * xd2;
            
            }
            if (textBox25.Text != "" && textBox23.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox25.Text);//#num
                xd2 = Convert.ToInt32(textBox23.Text);//price
                tottal += xd1 * xd2;
             
            }
            if (textBox26.Text != "" && textBox24.Text != "")
            {
                int xd1, xd2;
                xd1 = Convert.ToInt32(textBox26.Text);//#num
                xd2 = Convert.ToInt32(textBox24.Text);//price
                tottal += xd1 * xd2;
               
            }
            textBox39.Text = "";
            textBox39.Text = tottal.ToString();

        }

        private void button17_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel1.Visible = true;
        }

        private void button16_Click(object sender, EventArgs e)
        { textBox40.Text = "";
            int sum = 0,counter=1;
            
            int sum_10_bill = 0;
            for(int i = 0; i < info_today.bills_of_today.Count; i++)
            { int sum_bill = 0;
                for (int j = 0; j < info_today.bills_of_today[i].orders.Count; j++)
                {sum_bill+= info_today.bills_of_today[i].orders[j].its_price;
                    sum += info_today.bills_of_today[i].orders[j].its_price;
                }
                sum_10_bill += sum_bill;
                if (i == info_today.bills_of_today.Count - 1)
                {
                    textBox40.Text += counter + " = " + sum_10_bill + ".\r\n";
                    sum_10_bill = 0;
                } else if (i+1 == counter&&counter%10==0)
                {
                  
                    textBox40.Text += (counter - 9) + " .. " + counter + " = "+sum_10_bill + ".\r\n";
                    sum_10_bill = 0;
                }
                    counter++;
            }
            textBox40.Text += "Total ="+sum+" .";
            panel1.Visible = false;
            panel4.Visible = true;
        }
        int index_print_one_bill = 0;
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
               
                textBox42.Text = "";
                dataGridView2.Rows.Clear();
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                int id_order =Convert.ToInt32( row.Cells[0].Value);
                for (int i = 0; i < info_today.bills_of_today.Count; i++)
                {
                    if(info_today.bills_of_today[i].num_this_bill==id_order)
                    {
                        id_order = i;
                        break;
                    }
                }
               
                index_print_one_bill=id_order;
             textBox42.Text=   info_today.bills_of_today[id_order].name_cashier;
                textBox41.Text = info_today.bills_of_today[id_order].dat.ToString();
                textBox44.Text = info_today.bills_of_today[id_order].num_this_bill.ToString();
                int sum = 0;
                for(int j = 0; j < info_today.bills_of_today[id_order].orders.Count; j++)
                {
                    int n = dataGridView2.Rows.Add();
                    dataGridView2.Rows[n].Cells[0].Value = j + 1;
                    dataGridView2.Rows[n].Cells[1].Value = info_today.bills_of_today[id_order].orders[j].type_order;
                    dataGridView2.Rows[n].Cells[2].Value = info_today.bills_of_today[id_order].orders[j].its_price;
                    sum += info_today.bills_of_today[id_order].orders[j].its_price;
                }
                textBox43.Text = sum.ToString();
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            textBox43.Text = "";
            textBox42.Text = "";
            textBox41.Text = "";
            textBox44.Text = "";
            dataGridView1.Rows.Clear();
            for (int i = 0; i < info_today.bills_of_today.Count; i++)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = info_today.bills_of_today[i].num_this_bill;
                
            }
            panel6.Visible = true;
            panel1.Visible = false;
        }

        private void button23_Click_1(object sender, EventArgs e)
        {
            index_print_one_bill = -1;
            panel6.Visible = false;
            panel1.Visible = true;
        }

        private void listView10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private async void button24_Click(object sender, EventArgs e)
        {
            if (index_print_one_bill >= 0)
            {
             

                await Task.Run(() => print_bill_client_without_savedb(info_today.bills_of_today[index_print_one_bill]));

            }
        }
    }

    public static class StringUtil
    {
        private static byte[] key = new byte[8] { 1, 2, 3, 4, 5, 6, 7, 8 };
        private static byte[] iv = new byte[8] { 1, 2, 3, 4, 5, 6, 7, 8 };

        public static string Crypt(this string text)
        {
            SymmetricAlgorithm algorithm = DES.Create();
            ICryptoTransform transform = algorithm.CreateEncryptor(key, iv);
            byte[] inputbuffer = Encoding.Unicode.GetBytes(text);
            byte[] outputBuffer = transform.TransformFinalBlock(inputbuffer, 0, inputbuffer.Length);
            return Convert.ToBase64String(outputBuffer);
        }

        public static string Decrypt(this string text)
        {
            SymmetricAlgorithm algorithm = DES.Create();
            ICryptoTransform transform = algorithm.CreateDecryptor(key, iv);
            byte[] inputbuffer = Convert.FromBase64String(text);
            byte[] outputBuffer = transform.TransformFinalBlock(inputbuffer, 0, inputbuffer.Length);
            return Encoding.Unicode.GetString(outputBuffer);
        }
    }
}
