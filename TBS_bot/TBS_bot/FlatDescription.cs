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
        public string Address { get; }
        public string AlternationAddress { get; set; }
        public int RoomsCount { get; set; }
        public double flatArea { get; set; }
        public bool isAneks { get; set; }
        public bool isSend { get; set; }
        public string Link { get; set; }

        public string DetailedDescription { get; set; }

        public FlatDescription(string Address, string Link)
        {
            this.Address = Address;
            this.Link = Link;
        }

        public void FlatDescriptionUpdate(string number, string address, int flatNumber, double flatArea, bool isAneks, bool isSend, string DetailedDescription)
        {
            this.number = number;
            this.AlternationAddress = address;
            this.RoomsCount = flatNumber;
            this.flatArea = flatArea;
            this.isAneks = isAneks;
            this.isSend = isSend;
            this.DetailedDescription = DetailedDescription;
        }

        public string GetDetailedDescription()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("ogłosznie nr: " + number + Environment.NewLine);
            stringBuilder.Append(Address + Environment.NewLine);
            stringBuilder.Append(DetailedDescription + Environment.NewLine);
            stringBuilder.Append("powierzchnia: " + flatArea + Environment.NewLine);
            stringBuilder.Append("partycyp: " + (flatArea * 1200).ToString("F") + Environment.NewLine);
            stringBuilder.Append("czynsz: " + (flatArea * 14.25).ToString("F") + Environment.NewLine);
            stringBuilder.Append(Environment.NewLine);
            stringBuilder.Append("pokoje: " + RoomsCount + Environment.NewLine);
            stringBuilder.Append("aneks: " + isAneks);


            return stringBuilder.ToString();
        }

    }
}
