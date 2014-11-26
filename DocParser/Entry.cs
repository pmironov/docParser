using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToExcel
{
    class Entry
    {
        List<string> fPeople = new List<string>();

        internal string DocNumber { get; set; }
        internal string DocText { get; set; }
        internal string DocDate { get; set; }
        internal string ActNumber { get; set; }

        internal string FIO { get; set; }
        internal List<string> People { get { if (fPeople == null) return new List<string>(); else return fPeople; } set { fPeople = value; } }
        internal int RoomNum { get; set; }

        internal double SFull { get; set; }
        internal double SLiving { get; set; }

        internal string Address { get; set; }
        internal string Passport { get; set; }
    }
}
