using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx.convert.yrl
{
    public class OfferSaleNew
    {
        public OfferSaleNew(SalesAgent agent, int id, OfferType offerType, PropertyType propertyType, CategoryType categoryType, DateTime creationDate, DateTime lastUpdateDate,
            OfferLocation location, string url, DealStatusType dealStatus, decimal price)
        {
            Id = id;
            Type = offerType;
            Property = propertyType;
            Category = categoryType;
            Url = url;
            Agent = agent;
            CreationDate = creationDate;
            LastUpdateDate = lastUpdateDate;
            LocationPoint = location;
            DealStatus = dealStatus;
            Price = price;
        }

        public DealStatusType DealStatus { get; set; }

        public SalesAgent Agent { get; private set; }

        public OfferLocation LocationPoint { get; set; }

        public CategoryType Category { get; set; }

        public int Id { get; private set; }

        public OfferType Type { get; private set; }

        public PropertyType Property { get; set; }

        public DateTime CreationDate { get; set; }

        public DateTime LastUpdateDate { get; set; }

        public string Url { get; set; }

        public decimal Price { get; set; }


        public enum OfferType
        {
            [DescriptionAttribute("продажа")]
            Sale,
            [DescriptionAttribute("аренда")]
            Rent
        }

        public enum DealStatusType
        {
            [DescriptionAttribute("продажа от застройщика")]
            SaleDeveloper,
            [DescriptionAttribute("переуступка")]
            Reassignment
        }

        public enum PropertyType
        {
            [DescriptionAttribute("жилая")]
            Living,
            [DescriptionAttribute("коммерческая")]
            Commercial
        }

        public enum CategoryType
        {
            [DescriptionAttribute("дом")]
            House,
            [DescriptionAttribute("квартира")]
            Flat,
            [DescriptionAttribute("таунхаус")]
            Townhouse
        }

//        public static string GetDescription(Enum enumVal)
//        {
//            var attr = EnumsHelper.GetAttributeOfType<DescriptionAttribute>(enumVal);
//            return attr != null ? attr.Text : string.Empty;
//        }
    }
}
