using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TBS_bot
{
    class FlatDescription
    {
        public string number { get; set; }
        public string Addres { get; set; }
        public int flatNymbers { get; set; }
        public string flatArea { get; set; }
        public bool isAneks { get; set; }
        public bool isSend { get; set; }

        public FlatDescription(string number, string address, int flatNumber, double flatArea, bool isAneks, bool isSend)
        {
            this.number = number;
            this.Addres = address;
            this.flatNymbers = flatNumber;
            this.flatArea = flatArea.ToString("F");
            this.isAneks = isAneks;
            this.isSend = isSend;
        }
    }
}
