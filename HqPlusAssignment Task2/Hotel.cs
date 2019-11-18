using System;
using System.Collections.Generic;
using System.Text;

namespace HqPlusAssignment_Task2
{
    public class Hotel
    {
        public int hotelID { get; set; }

        public string BookingPageUrl { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string Description { get; set; }
        public int? Classification { get; set; }
        public float? ReviewScore { get; set; }
        public int? NumberOfReviews { get; set; }
        public List<string> RoomCategories { get; set; }
        public List<Hotel> AlternativeHotels { get; set; }
    }
}
