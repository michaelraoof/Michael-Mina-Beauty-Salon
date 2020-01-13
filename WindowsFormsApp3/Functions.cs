using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography;
using Ionic.Zip;
using System.Threading;
namespace WindowsFormsApp3
{
    class Functions
    {
        public void WriteInFileFromList_attendence(List< attendence> E, string file_name,int index)
        {
            
            FileStream fs = new FileStream(Environment.CurrentDirectory + "/att/" + file_name + ".txt", FileMode.Create);
            StreamWriter sr = new StreamWriter(fs);
           // string[] feild = file_name.Split('-');
           
            for (int i = 0; i < E[index].employees_month.Count; i++)
            {
                string rec = E[index].employees_month[i].name;
                for (int j = 0; j < E[index].employees_month[i].time_working.Count; j++)
                {
                    if ( E[index].employees_month[i].time_working[j].name_cashier_end == null )
                    {
                        rec = rec + "*no*no*" + E[index].employees_month[i].time_working[j].name_cashier_start + '*' + E[index].employees_month[i].time_working[j].startime.ToString();
                    
                    }
                    else
                    {


                        rec = rec + '*' + E[index].employees_month[i].time_working[j].endtime.ToString() + '*' + E[index].employees_month[i].time_working[j].name_cashier_end + '*' + E[index].employees_month[i].time_working[j].name_cashier_start + '*' + E[index].employees_month[i].time_working[j].startime.ToString();
                    }
                }
                rec = rec.Crypt();
                sr.WriteLine(rec);
            }
            sr.Close();
        }
        public void fillList_attendence(ref List< attendence> E,string file_name)
        {
            if (!File.Exists(Environment.CurrentDirectory + "/att/" + file_name + ".txt"))
            {
                FileStream fs = new FileStream(Environment.CurrentDirectory + "/att/" + file_name + ".txt", FileMode.Create);

            }
            else
            {
                attendence gr = new attendence();
                FileStream fs = new FileStream(Environment.CurrentDirectory + "/att/" + file_name + ".txt", FileMode.Open);
                StreamReader sr = new StreamReader(fs);
               
                while (sr.Peek() != -1)
                {
                    string rec = sr.ReadLine();
                    rec = rec.Decrypt();
                    string[] feild = rec.Split('*');
                    onebyone obj = new onebyone();
                    obj.name = feild[0];
                 
                    for (int i = 1; i < feild.Length; i += 4)
                    {
                        time_all_months te = new time_all_months();
                        if (feild[i] == "no" || (feild[i + 1]) == "no")
                        {

                        }
                        else
                        {
                            te.endtime = Convert.ToDateTime(feild[i]);

                            te.name_cashier_end = (feild[i + 1]);
                        }
                      te.name_cashier_start= (feild[i + 2]);
                     
                      te.startime= Convert.ToDateTime(feild[i + 3]);
                     
                    
                        obj.time_working.Add(te);
                    }
                    gr.employees_month.Add(obj);
                    
                }
                E.Add(gr);
                sr.Close();
            }
        }//done


        //fill the list
        public void fillList_cashier(List<cashier> E,ref today oo)
        {
            if (!File.Exists("asd.txt"))
            {
                FileStream fs = new FileStream("asd.txt", FileMode.Create);
                oo.name_admin = "admin";
                oo.password_admin = "admin";
            }
            else
            {
                FileStream fs = new FileStream("asd.txt", FileMode.Open);
                StreamReader sr = new StreamReader(fs);
                string rec = sr.ReadLine();
                rec = rec.Decrypt();
                oo.name_admin = rec;
                rec = sr.ReadLine();
                rec = rec.Decrypt();
                oo.password_admin = rec;
                while (sr.Peek() != -1)
                {
                     rec = sr.ReadLine();
                    rec = rec.Decrypt();
                    string[] feilds = rec.Split('*');
                    cashier temp = new cashier();
                    temp.name = feilds[0];
                    temp.password = feilds[1];

                    E.Add(temp);
                }
                sr.Close();
            }
        }//done



        public void fillList_employee(List<employee> E)
        {
            if (!File.Exists("abc.txt"))
            {

                FileStream fs = new FileStream("abc.txt", FileMode.Create);
            }
            else
            {
                FileStream fs = new FileStream("abc.txt", FileMode.Open);
                StreamReader sr = new StreamReader(fs);

                while (sr.Peek() != -1)
                {
                    string rec = sr.ReadLine();
                    rec = rec.Decrypt();
                    string[] feilds = rec.Split('*');
                    employee temp = new employee();
                                                  
                    temp.id = feilds[0];
                    temp.name = feilds[1];
                    temp.num_days = Convert.ToInt32(feilds[2]);
                  
                    temp.time_start_work = Convert.ToDateTime(feilds[3]);
                    temp.address = feilds[4];
                    //
                   
                    temp.sum_over_salary = Convert.ToInt32(feilds[5]);
                    temp.sum_subtraction_salary = Convert.ToInt32(feilds[6]);
                    rec = sr.ReadLine();
                    rec = rec.Decrypt();
                    if (Convert.ToInt32(rec) != 0)
                    {
                        for (int i = 0; i < Convert.ToInt32(rec); i++)
                        {
                            rec = sr.ReadLine();
                            rec = rec.Decrypt();
                            string[] feilds_2 = rec.Split('*');
                            over_salary obj = new over_salary();
                            obj.amount_over_salary = Convert.ToInt32(feilds_2[0]);
                            obj.describtion_over_salary = feilds_2[1];
                            obj.time_over_salary = Convert.ToDateTime(feilds_2[2]);
                            temp.over_salry.Add(obj);
                        }
                    }
                    rec = sr.ReadLine();
                    rec = rec.Decrypt();
                    if (Convert.ToInt32(rec) != 0)
                    {
                        for (int i = 0; i < Convert.ToInt32(rec); i++)
                        {
                            rec = sr.ReadLine();
                            rec = rec.Decrypt();
                            string[] feilds_2 = rec.Split('*');
                            subtraction_salary obj = new subtraction_salary();
                            obj.amount_subtraction_salary = Convert.ToInt32(feilds_2[0]);
                            obj.describtion_subtraction_salary = feilds_2[1];
                            obj.time_subtraction_salary = Convert.ToDateTime(feilds_2[2]);
                            temp.subtraction_salry.Add(obj);
                        }
                    }

                    E.Add(temp);
                }
                sr.Close();
            }
        }
     
    

        // After adding all events , write in file using the list not in use currently
        public void WriteInFileFromList_cashier(List<cashier> a,ref today oo)
        {
            FileStream fs = new FileStream("asd.txt", FileMode.Create);
            StreamWriter sr = new StreamWriter(fs);
            string rec = oo.name_admin;
            rec = rec.Crypt();
            sr.WriteLine(rec);
            rec = oo.password_admin;
            rec = rec.Crypt();
            sr.WriteLine(rec);
            for (int i = 0; i < a.Count(); i++)
            {
                 rec = a[i].name + '*' + a[i].password;
                rec = rec.Crypt();
                sr.WriteLine(rec);
            }
            sr.Close();
        }

        // After adding all events , write in file using the del list not in use currently
        public void WriteInFileFromList_employee(List<employee> a)
        {
            FileStream fs = new FileStream("abc.txt", FileMode.Create);
            StreamWriter sr = new StreamWriter(fs);

            for (int i = 0; i < a.Count(); i++)
            {
                string rec = a[i].id + '*' + a[i].name + '*' + a[i].num_days + '*' + a[i].time_start_work + '*' + a[i].address + '*' + a[i].sum_over_salary + '*' + a[i].sum_subtraction_salary;
                rec = rec.Crypt();
                sr.WriteLine(rec);
                rec = a[i].over_salry.Count.ToString();
                rec = rec.Crypt();
                sr.WriteLine(rec);
                for(int jj = 0; jj < a[i].over_salry.Count; jj++)
                {
                    rec = a[i].over_salry[jj].amount_over_salary.ToString() + '*' + a[i].over_salry[jj].describtion_over_salary + '*' + a[i].over_salry[jj].time_over_salary.ToString();
                    rec = rec.Crypt();
                    sr.WriteLine(rec);
                }
                rec = a[i].subtraction_salry.Count.ToString();
                rec = rec.Crypt();
                sr.WriteLine(rec);
                for (int jj = 0; jj < a[i].subtraction_salry.Count; jj++)
                {
                    rec = a[i].subtraction_salry[jj].amount_subtraction_salary.ToString() + '*' + a[i].subtraction_salry[jj].describtion_subtraction_salary + '*' + a[i].subtraction_salry[jj].time_subtraction_salary.ToString();
                    rec = rec.Crypt();
                    sr.WriteLine(rec);
                }

            }
            sr.Close();
        }
        public void fillList_info_today(ref today E, List<employee> employees)
        {
            if (!File.Exists("hestory.txt"))
            {

                FileStream fs = new FileStream("hestory.txt", FileMode.Create);
            }
            else
            {
                FileStream fs = new FileStream("hestory.txt", FileMode.Open);
                StreamReader sr = new StreamReader(fs);

                while (sr.Peek() != -1)
                {
                    string rec = sr.ReadLine();
                 //   rec = rec.Decrypt();
                    string[] feilds = rec.Split('*');
                    today temp = new today();
                    
                    temp.num_bills = Convert.ToInt32(feilds[0]);
                    temp.num_employees = Convert.ToInt32(feilds[1]);
                    temp.time_today = Convert.ToDateTime(feilds[2]);

                    if (temp.num_employees != 0)
                    {

                        for (int i = 0; i < temp.num_employees; i++)
                        {
                            rec = sr.ReadLine();
                          //  rec = rec.Decrypt();
                            string[] ffeilds = rec.Split('*');

                            for (int j = 0; j < employees.Count; j++)
                            {
                                if (employees[j].name == feilds[0])
                                {

                                    employees[j].time_coming_today.startime = Convert.ToDateTime(ffeilds[1]);
                                    employees[j].time_coming_today.name_cashier = (ffeilds[2]);
                                    if (ffeilds[3] != "no" && ffeilds[4] != "no")
                                    {
                                        employees[j].time_leaving_today.endtime = Convert.ToDateTime(ffeilds[3]);
                                        employees[j].time_leaving_today.name_cashier = (ffeilds[4]);
                                    }
                                }
                            }
                            temp.today_emloyees.Add(feilds[0]);
                        }
                    }int hh = 0;
                    for (int i = 0; i < temp.num_bills; i++)
                    {
                        rec = sr.ReadLine();
                       // rec = rec.Decrypt();
                        string[] feild = rec.Split('*');
                        bill ui = new bill();
                        ui.name_cashier= feild[0];
                        
                        ui.num_this_bill = Convert.ToInt32(feild[1]);
                        hh = ui.num_this_bill;
                        for (int j = 0; j < Convert.ToInt32(feild[2]); j++)
                        {
                            one_order solo_order = new one_order();
                            solo_order.its_price = Convert.ToInt32(feild[3 + (j * 2)]);
                            solo_order.type_order = feild[4 + (j * 2)];
                          ui.orders.Add(solo_order);
                           
                        }
                temp.bills_of_today.Add(ui);
                       
                    }
                    temp.num_bills = hh;
                    E = temp;
                }
                sr.Close();
            }
        }
        public void WriteInFileFromList_info_today(today temp, List<employee> employees)
        {
            FileStream fs = new FileStream("hestory.txt", FileMode.Create);
            StreamWriter sr = new StreamWriter(fs);
            string rec;
            rec = temp.bills_of_today.Count.ToString() + '*' + temp.num_employees.ToString()+'*'+ temp.time_today.ToString();
           // rec = rec.Crypt();
            sr.WriteLine(rec);
            for (int i = 0; i < temp.num_employees; i++)
            {
                for(int j = 0; j < employees.Count; j++)
                {
                    if (temp.today_emloyees[i] == employees[j].name)
                    {
                        rec = employees[j].name + '*' + employees[j].time_coming_today.startime.ToString() + '*' + employees[j].time_coming_today.name_cashier;
                        if (employees[j].time_leaving_today.name_cashier == null && employees[j].time_leaving_today.endtime == null)
                        {
                            rec += '*' + "no" + '*' + "no";
                        }
                        else
                        {
                            rec += '*' + employees[j].time_leaving_today.endtime.ToString() + '*' + employees[j].time_leaving_today.name_cashier;
                        }
                    //   rec = rec.Crypt();
                        sr.WriteLine(rec);

                    }

                }
              
            }
            for(int i = 0; i < temp.num_bills; i++)
            {
                rec = temp.bills_of_today[i].name_cashier + '*' + temp.bills_of_today[i].num_this_bill.ToString() + '*' + temp.bills_of_today[i].orders.Count;
                    for (int j = 0; j < temp.bills_of_today[i].orders.Count; j++)
                    {
                    rec = rec + '*' + temp.bills_of_today[i].orders[j].its_price.ToString() + '*' + temp.bills_of_today[i].orders[j].type_order;
                    }
           //    rec = rec.Crypt();
                sr.WriteLine(rec);
            }

            sr.Close();
        }
      
        
        public void compress_files(string password)
        {
            using (ZipFile zip = new ZipFile())
            {
                ZipEntry e1 = new ZipEntry();

                if (System.IO.File.Exists(Environment.CurrentDirectory + "/abc.txt"))
                {

                    e1 = zip.AddFile("abc.txt");
                    e1.Password = password;
                    e1.Encryption = EncryptionAlgorithm.WinZipAes256;
                }
                if (System.IO.File.Exists(Environment.CurrentDirectory + "/asd.txt"))
                {
                    e1 = zip.AddFile("asd.txt");
                    e1.Password = password;
                    e1.Encryption = EncryptionAlgorithm.WinZipAes256;
                }
                if (System.IO.File.Exists(Environment.CurrentDirectory + "/hestory.txt"))
                {
                    e1 = zip.AddFile("hestory.txt");
                    e1.Password = password;
                    e1.Encryption = EncryptionAlgorithm.WinZipAes256;
                }
                
                zip.Save("new.zip");
                


            }

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();

            if (System.IO.File.Exists(Environment.CurrentDirectory + "/abc.txt"))
            {
                System.IO.File.Delete(Environment.CurrentDirectory + "/abc.txt");
            }
            if (System.IO.File.Exists(Environment.CurrentDirectory + "/asd.txt"))
            {
                System.IO.File.Delete(Environment.CurrentDirectory + "/asd.txt");
            }
            if (System.IO.File.Exists(Environment.CurrentDirectory + "/hestory.txt"))
            {
                System.IO.File.Delete(Environment.CurrentDirectory + "/hestory.txt");
            }
        }
        public void extract_files(string password)
        {
            if (!System.IO.File.Exists(Environment.CurrentDirectory + "/new.zip"))
            {
                FileStream fs = new FileStream(Environment.CurrentDirectory + "asd.txt", FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                string rrr = "admin";
                rrr = rrr.Crypt();
                sw.WriteLine(rrr);
                sw.WriteLine(rrr);
                sw.Close();
                FileStream fs1 = new FileStream(Environment.CurrentDirectory + "abc.txt", FileMode.Create);
                StreamReader sr = new StreamReader(fs1);
                sr.Close();
            }
            else
            {
                using (ZipFile zip = new ZipFile("new.zip"))
                {

                    zip.Password = (password);
                    zip.Encryption = EncryptionAlgorithm.WinZipAes256;
                    zip.StatusMessageTextWriter = Console.Out;

                    zip.ExtractAll(Environment.CurrentDirectory, ExtractExistingFileAction.Throw);
                }
            }
            if (System.IO.File.Exists(Environment.CurrentDirectory + "/new.zip"))
            {
                System.IO.File.Delete(Environment.CurrentDirectory + "/new.zip");
            }
           
            }
        
        //check if done or not
       /* public void Donecheck(List<Event> events, List<Event> afterdel)
        {
            for (int i = 0; i < events.Count(); i++)
            {
                if (events[i].dateandtimeofend < DateTime.Now)
                {
                    events[i].done = true;
                    afterdel.Add(events[i]);
                    events.Remove(events[i]);
                }
            }
        }

        //del event
        public void deleteEvent(List<Event> events, List<Event> del, DateTime find)
        {
            for (int i = 0; i < events.Count(); i++)
            {
                if (events[i].dateandtimeofstart.Year == find.Year && events[i].dateandtimeofstart.Day == find.Day && events[i].dateandtimeofstart.Month == find.Month)
                {
                    del.Add(events[i]);                                                                           
                    events.Remove(events[i]);
                }
            }
        }

        //reminder viewer
        public string showReminder(List<Event> events)
        {
            string h = "No event tomorrow";
            for (int i = 0; i < events.Count; i++)
            {
                if (events[i].rem.reminderTime.Year == DateTime.Now.Year && events[i].rem.reminderTime.Month == DateTime.Now.Month && events[i].rem.reminderTime.Day == DateTime.Now.Day)
                {
                    h = "You have an Event at : " + events[i].dateandtimeofstart + " Your Reminder Message is : " + events[i].rem.remind;
                }
            }

            return h;
        }

    */

    }
}
