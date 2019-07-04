using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DataTables
{
    class CompanyTable
    {
        public int Id { get; set; }
        public int Pid { get; set; }
        public int Sid { get; set; }
        public string Name { get; set; }
        public bool Locked { get; set; }
        public int Id_user_create { get; set; }
        public DateTime Date_create { get; set; }
        public int Id_user_modified { get; set; }
        public DateTime Date_modified { get; set; }
        public string Xdoc { get; set; }
        public int Node { get; set; }
        public string Profile { get; set; }
    }
}
