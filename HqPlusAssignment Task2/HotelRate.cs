using System;
using System.Collections.Generic;
using System.Text;

namespace HqPlusAssignment_Task2
{
    public class HotelRate
    {
        public DateTime ArrivalDate { get; set; }
        public DateTime DepartureDate { get; set; }
        public float Price { get; set; }
        public int PriceNumericInteger { get; set; }        
        public string Currency { get; set; }
        public string RateName { get; set; }
        public int Adults { get; set; }
        public bool BreakfastIncluded { get; set; }
    }
}
