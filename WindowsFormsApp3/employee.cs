using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp3
{
    class employee
    {
        public string id;
        public string address;
        public DateTime time_start_work = new DateTime();
        public string name ;
       
       
        public int num_days;
       
       
        public List<over_salary> over_salry = new List<over_salary>();
        public List<subtraction_salary> subtraction_salry = new List<subtraction_salary>();

        public int sum_over_salary;
        public int sum_subtraction_salary;

        public time_coming time_coming_today = new time_coming();
        public time_leaving time_leaving_today = new time_leaving();
       
    }
   
    class time_coming
    {
        public DateTime startime = new DateTime();
        public string name_cashier;
    }
    class time_leaving
    {
        public DateTime endtime = new DateTime();
        public string name_cashier;
    }
    class over_salary
    {
        public int amount_over_salary;
        public string describtion_over_salary;
        public DateTime time_over_salary=new DateTime();
        
    }
    class subtraction_salary
    {
        public int amount_subtraction_salary;
        public string describtion_subtraction_salary;
        public DateTime time_subtraction_salary = new DateTime();
    }
}
