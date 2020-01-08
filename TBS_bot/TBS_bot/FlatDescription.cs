using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TBS_bot
{
    class FlatDescription
    {
        public string Number { get; set; }
        public string Address { get; }
        public string AlternationAddress { get; set; }
        public int RoomsCount { get; set; }
        public double FlatArea { get; set; }
        public bool IsAneks { get; set; }
        public bool IsSend { get; set; }
        public string Link { get; set; }
        public string District { get; set; }
        public string DetailedDescription { get; set; }
        public bool IsClassified { get; set; }

        public double Participation { get; set; }

        public FlatDescription(string Address, string Link)
        {
            this.Address = Address;
            this.Link = Link;
        }

        public void FlatDescriptionUpdate(string number, string address, int flatNumber, double flatArea, bool isAneks, string district, string detailedDescription, double participation)
        {
            this.Number = number;
            this.AlternationAddress = address;
            this.RoomsCount = flatNumber;
            this.FlatArea = flatArea;
            this.IsAneks = isAneks;
            this.District = district;
            this.DetailedDescription = detailedDescription;
            this.Participation = participation;
            this.IsSend = false;
        }

        public string GetDetailedDescription()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append("ogłosznie nr: " + Number + Environment.NewLine);
            stringBuilder.Append(Address + Environment.NewLine);
            stringBuilder.Append("osiedle: " + District + Environment.NewLine);
            stringBuilder.Append(DetailedDescription + Environment.NewLine);
            stringBuilder.Append("powierzchnia: " + FlatArea + Environment.NewLine);
            stringBuilder.Append("partycyp: " + Participation + Environment.NewLine);
            stringBuilder.Append("metr: " + (Participation / FlatArea).ToString("F") + Environment.NewLine);
            stringBuilder.Append("czynsz: " + (FlatArea * 14.25).ToString("F") + Environment.NewLine);
            stringBuilder.Append(Environment.NewLine);
            stringBuilder.Append("pokoje: " + RoomsCount + Environment.NewLine);
            stringBuilder.Append("aneks: " + IsAneks + Environment.NewLine);
            stringBuilder.Append("zakwalifikowany: " + IsClassified);


            return stringBuilder.ToString();
        }

    }
}
