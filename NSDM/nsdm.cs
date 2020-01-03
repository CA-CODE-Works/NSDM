using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace NSDM
{
    class nsdm
    {
        [DataContract]
        public class DynamicAttribute
        {
            [DataMember (EmitDefaultValue = false)]
            public string groupName { get; set; }
            [DataMember]
            public string attributeName { get; set; }
            [DataMember (EmitDefaultValue = false)]
            public string attributeValue { get; set; }
        }

        [DataContract]
        public class ShelfObject
        {
            [DataMember (Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember(Name = "$action")]
            public string action { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string vendor { get; set; }
            [DataMember]
            public string model { get; set; }
            [DataMember]
            public string rev { get; set; }
            [DataMember]
            public string networkId { get; set; }
            [DataMember]
            public string dimensionUnits { get; set; }
            [DataMember]
            public string barCode { get; set; }
            [DataMember]
            public string batchNumber { get; set; }
            [DataMember]
            public string purchaseDate { get; set; }
            [DataMember]
            public string purchasePrice { get; set; }
            [DataMember]
            public string serialNumber { get; set; }
            [DataMember]
            public string orderNum { get; set; }
            [DataMember]
            public string installedDate { get; set; }
            [DataMember]
            public string comments { get; set; }
            [DataMember]
            public string targetId { get; set; }
            [DataMember]
            public string device_id { get; set; }
            [DataMember(Name = "$parentSite")]
            public List<string> parentSite { get; set; }
            [DataMember(Name = "$ServiceInstance")]
            public List<string> ServiceInstance { get; set; }
            [DataMember(EmitDefaultValue = false)]
            public List<DynamicAttribute> dynamicAttributes { get; set; }
        }
        [DataContract]
        public class PortObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string bandWidth { get; set; }
            [DataMember]
            public string direction { get; set; }
            [DataMember]
            public string portAccessId { get; set; }
            [DataMember]
            public string connectorType { get; set; }
            [DataMember]
            public string channelization { get; set; }
            [DataMember]
            public string parentPortChanName { get; set; }
            [DataMember]
            public string networkId { get; set; }
            [DataMember]
            public string description { get; set; }
            [DataMember]
            public string role { get; set; }
            [DataMember]
            public string hecig { get; set; }
            [DataMember]
            public string networkDirection { get; set; }
            [DataMember]
            public string physicalPortName { get; set; }
            [DataMember]
            public string aidFormula { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember(Name = "$site")]
            public string site { get; set; }
            [DataMember(Name = "$portClass")]
            public string portClass { get; set; }
            [DataMember(Name = "$parentShelf")]
            public List<string> parentShelf { get; set; }
        }

        [DataContract]
        public class BillingObject
        {
            [DataMember]
            public BillingCustObject billingCust { get; set; }
            [DataMember]
            public BillingServiceObject billingService { get; set; }
            [DataMember]
            public BillingDataObject billingData { get; set; }
        }


        [DataContract]
        public class BillingDataObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
            [DataMember]
            public string datasource { get; set; }
            [DataMember]
            public string srNo { get; set; }
            [DataMember]
            public string compCode { get; set; }
            [DataMember]
            public string resourceType1 { get; set; }
            [DataMember]
            public string shiftCode { get; set; }
            [DataMember]
            public string periodicity { get; set; }
            [DataMember]
            public string sharingPer { get; set; }
            [DataMember]
            public string effectiveDate { get; set; }
            [DataMember]
            public string deInstallDate { get; set; }
            [DataMember]
            public string chargeType { get; set; }
            [DataMember]
            public string chargeCategory { get; set; }
            [DataMember]
            public string billAction { get; set; }
            [DataMember]
            public string qty { get; set; }
            [DataMember]
            public string rate { get; set; }
            [DataMember]
            public string calnetProductId { get; set; }
            [DataMember]
            public string vendorAccount { get; set; }
            [DataMember]
            public string circuitId { get; set; }
            [DataMember]
            public string deviceName { get; set; }
            [DataMember]
            public string deviceType { get; set; }
            [DataMember]
            public string deviceSerialNum { get; set; }
            [DataMember]
            public string deviceModel { get; set; }
            [DataMember]
            public string phoneNumber { get; set; }
            [DataMember]
            public string recordName { get; set; }
            [DataMember]
            //public string lastModifiedDate { get; set; }
            public string lastUpdatedBy { get; set; }
            [DataMember]
            public string previousSrNo { get; set; }
            [DataMember]
            public string comments { get; set; }
            [DataMember]
            public string commentsTemp { get; set; }
            [DataMember]
            public string billExpiryDate { get; set; }
            [DataMember]
            public string rateDisplay { get; set; }
            //[DataMember]
            //public string rawRate { get; set; }
            [DataMember]
            public string isRateUpdate { get; set; }
            [DataMember(Name = "$customer")]
            public List<string> customer { get; set; }
            [DataMember(Name = "$ServiceInstance")]
            public List<string> ServiceInstance { get; set; }
        }
        [DataContract]
        public class BillingCustObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
        }
        [DataContract]
        public class PortTypeObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember(Name = "$parentSite")]
            public List<string> parentSite { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
            [DataMember]
            public string name { get; set; }

        }
        //public class PortTypeObject2
        //{
        //    [DataMember(Name = "$type")]
        //    public string type1 { get; set; }
        //    [DataMember(Name = "$refId")]
        //    public string refId { get; set; }
        //    [DataMember(Name = "$parentSite")]
        //    public List<string> parentSite { get; set; }
        //    [DataMember(Name = "$action")]
        //    public string action1 { get; set; }
        //    [DataMember]
        //    public string name { get; set; }

        //}

        [DataContract]
        public class BillingServiceObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
        }


        [DataContract]
        public class ChannelObject
        {
            [DataMember(Name = "name")]
            public string name { get; set; }
            [DataMember]
            public string band_width { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string parent_trail { get; set; }
            [DataMember]
            public string child_trail { get; set; }
        }

        [DataContract]
        public class SIOtherObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
        }

        [DataContract]
        public class SIShelfObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember(Name = "$parentSite")]
            public List<string> parentSite { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
        }

        [DataContract]
        public class SiteObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember(Name = "name")]
            public string name { get; set; }
            [DataMember(Name = "siteId")]
            public string siteID { get; set; }
            [DataMember(Name = "type")]
            public string type { get; set; }
            [DataMember(Name = "status")]
            public string status { get; set; }
            [DataMember(Name = "clli", EmitDefaultValue = false)]
            public string clli { get; set; }
            //[DataMember(Name = "LATA", EmitDefaultValue = false)]
            //public string LATA { get; set; }
            [DataMember(Name = "address")]
            public string address { get; set; }
            [DataMember(Name = "city")]
            public string city { get; set; }
            [DataMember(Name = "stateProv")]
            public string stateProv { get; set; }
            [DataMember(Name = "postalCode1")]
            public string postalCode1 { get; set; }
            //[DataMember(Name = "floor")]
            //public string floor { get; set; }
            [DataMember(EmitDefaultValue = false)]
            public string room { get; set; }
            //[DataMember(Name = "inServiceDate")]
            //public string inServiceDate { get; set; }
            //[DataMember(Name = "comments")]
            //public string comments { get; set; }
            //[DataMember(Name = "$parentSiteName")]
            //public List<string> parentSiteName { get; set; }
            //[DataMember(Name = "$customer")]
            //public List<string> customer { get; set; }

            [DataMember(Name = "dynamicAttributes", EmitDefaultValue = false)]
            public List<DynamicAttribute> dynamicAttributes { get; set; }

        }
        [DataContract]
        public class NIObject
        {
            [DataMember]
            public string ni_type { get; set; }

            [DataMember(Name = "name")]
            public string name { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string parent_site_name { get; set; }
            [DataMember]
            public string parent_shelf_name { get; set; }
            [DataMember]
            public string port_name { get; set; }
        }
        [DataContract]
        public class LIObject
        {
            [DataMember(Name = "name")]
            public string name { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string installed_date { get; set; }
        }
        [DataContract]
        public class VRFObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string routeMapOut { get; set; }
            [DataMember]
            public string comments { get; set; }
        }
        [DataContract]
        public class SCGObject
        {
            [DataMember(Name = "name")]
            public string name { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
        }
        [DataContract]
        public class LocalContactObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string description { get; set; }
            [DataMember]
            public string phoneNumber { get; set; }
            [DataMember]
            public string alternatePhoneNumber { get; set; }
            [DataMember]
            public string emailAddress { get; set; }

            [DataMember(Name = "$site")]
            public List<string> site { get; set; }
        }
        [DataContract]
        public class CustomerObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string customerName { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string primaryPhoneNo { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string billingAddress { get; set; }
            [DataMember]
            public string city { get; set; }
            [DataMember]
            public string stateProv { get; set; }
            [DataMember]
            public string postalCode1 { get; set; }
            [DataMember]
            public string postalCode2 { get; set; }
            [DataMember]
            public string country { get; set; }
            [DataMember]
            public string npaNxx { get; set; }
            [DataMember]
            public string floor { get; set; }
            [DataMember]
            public string room { get; set; }
            [DataMember]
            public string billingCode { get; set; }
            [DataMember]
            public string billingFax { get; set; }
            [DataMember]
            public string comments { get; set; }
            [DataMember]
            public string billingContacts { get; set; }
            [DataMember]
            public string techContacts { get; set; }
        }
        [DataContract]
        public class ServiceInstanceObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string isparentServ { get; set; }
            [DataMember]
            public string serviceType { get; set; }
            [DataMember]
            public string serviceStatus { get; set; }
            [DataMember]
            public string description { get; set; }
            [DataMember]
            public string svcRequestNumber { get; set; }
            [DataMember]
            public string comments { get; set; }
            [DataMember]
            public string vendor { get; set; }
            [DataMember]
            public string CustAccountCode { get; set; }
            [DataMember]
            public string parentServiceName { get; set; }
            [DataMember]
            public string installedDate { get; set; }
            [DataMember]
            public string decommissionDate { get; set; }
            [DataMember]
            public string lanActivationDate { get; set; }
            [DataMember]
            public string srCompletionDate { get; set; }
            [DataMember]
            public string srDecommissionDate { get; set; }
            [DataMember(Name = "$site")]
            public List<string> site { get; set; }
            [DataMember(Name = "$customer")]
            public List<string> customer { get; set; }
            [DataMember(Name = "$trail")]
            public List<string> trail { get; set; }
            [DataMember(Name = "$bandwidth")]
            public List<string> bandwidth { get; set; }
        }
        [DataContract]
        public class STECustObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
        }
        [DataContract]
        public class STEObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember(Name = "$site")]
            public List<string> site { get; set; }
            [DataMember(Name = "$bandwidth")]
            public List<string> bandwidth { get; set; }
        }
        [DataContract]
        public class TrailObject
        {
            [DataMember(Name = "$type")]
            public string type2 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember(Name = "$action")]
            public string action1 { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string assignmentType { get; set; }
            [DataMember]
            public string bandWidth { get; set; }
            [DataMember]
            public string protectionType { get; set; }
            [DataMember]
            public string owner { get; set; }
            [DataMember]
            public string direction { get; set; }
            [DataMember]
            public string billingCode { get; set; }
            [DataMember]
            public string comments { get; set; }
            [DataMember]
            public string decommissionDate { get; set; }
            [DataMember]
            public string installedDate { get; set; }
            [DataMember(Name = "$zSideSite")]
            public List<string> zSideSite { get; set; }
            [DataMember(Name = "$aSideSite")]
            public List<string> aSideSite { get; set; }
            [DataMember(Name = "$ServiceInstance")]
            public List<string> ServiceInstance { get; set; }
            [DataMember(EmitDefaultValue = false)]
            public List<DynamicAttribute> dynamicAttributes { get; set; }
        }
        [DataContract]
        public class ChildTrailObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
        }
        [DataContract]
        public class TmsVlanObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string comments { get; set; }
        }
        [DataContract]
        public class LogicalInterfaceObject
        {
            [DataMember(Name = "$type")]
            public string type1 { get; set; }
            [DataMember(Name = "$refId")]
            public string refId { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string type { get; set; }
            [DataMember]
            public string status { get; set; }
            [DataMember]
            public string ceVlanidEvcMap { get; set; }
            [DataMember]
            public string localAsNumber { get; set; }
            [DataMember]
            public string routerIdIpAddress { get; set; }
            [DataMember (Name = "$vrfs")]
            public List<string> vrfs { get; set; }
            [DataMember(Name = "$networkInterface")]
            public List<string> networkInterface { get; set; }
        }
    }
}
