using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp3
{
    class attendence
    {
        public List<onebyone> employees_month = new List<onebyone>();
    }
    class onebyone {

        public string name;
        public List<time_all_months> time_working = new List<time_all_months>();

    }

    class time_all_months
    {
        public DateTime startime = new DateTime();
        public string name_cashier_start;
        public DateTime endtime = new DateTime();
        public string name_cashier_end;
    }
}
