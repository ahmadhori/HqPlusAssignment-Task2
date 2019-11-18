using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HqPlusAssignment_Task2
{
    public class JsonHelper
    {
        public static IList<HotelRate> ParseHotelRatesFile(string jsonText)
        {
            JObject jsonFile = JObject.Parse(jsonText);
            
            //Parsing Hotel
            //string hotelString = jsonFile["hotel"].ToString();
            //Hotel hotel = JsonConvert.DeserializeObject<Hotel>(hotelString);
            
            //Parsing Hotel Rates
            IList<JToken> results = jsonFile["hotelRates"].Children().ToList();
            IList<HotelRate> hotelRates = new List<HotelRate>();
            
            foreach (JToken result in results)
            {
                HotelRate hr = new HotelRate();

                //Parsing Adults Field
                int adults;
                if (Int32.TryParse(result["adults"].ToString(), out adults))
                {
                    hr.Adults = adults;
                }

                //Parsing ArrivalDate Field
                DateTime arrivalDate;
                if (DateTime.TryParse(result["targetDay"].ToString(), out arrivalDate))
                {
                    hr.ArrivalDate = arrivalDate;
                }

                //Parsing los and calculating DepartureDate
                int los;
                if (Int32.TryParse(result["los"].ToString(), out los))
                {
                    if (hr.ArrivalDate != null)
                    {
                        hr.DepartureDate = arrivalDate.AddDays(los);
                    }
                }

                //Parsing Price Fields
                float price;
                if (float.TryParse(result["price"]["numericFloat"].ToString(), out price))
                {
                    hr.Price = price;
                }

                int priceNumericInteger;
                if (int.TryParse(result["price"]["numericInteger"].ToString(), out priceNumericInteger))
                {
                    hr.PriceNumericInteger = priceNumericInteger;
                }
                hr.Currency = result["price"]["currency"].ToString();

                //Parsing RateName Field
                hr.RateName = result["rateName"].ToString();

                //Parsing BreakfastIncluded Field
                bool breakfast;
                if (bool.TryParse(result["rateTags"][0]["shape"].ToString(), out breakfast))
                {
                    hr.BreakfastIncluded = breakfast;
                }

                hotelRates.Add(hr);
            }
            return hotelRates;
        }
    }
}
