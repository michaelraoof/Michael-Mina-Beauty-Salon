using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp3
{
    class   today
    {
        public DateTime time_today;
     public   List<string> today_emloyees=new List<string>();
     public   int num_employees;
     public   int num_bills;
        public string password_admin;
        public string name_admin;
        public List<bill> bills_of_today = new List<bill>();
    }
    class one_order
    {
       
        public string type_order;
        public int its_price;
        public int id;
    }
    class bill
    {
      public  List<one_order> orders = new List<one_order>();
        public int num_this_bill;
        public string name_cashier;
        public DateTime dat = new DateTime();
    }
}
