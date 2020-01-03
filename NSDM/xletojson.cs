
using System;
using System.Linq;
using System.Data.OleDb;
using System.Data.Common;
using Newtonsoft.Json;
using System.IO;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace NSDM
{
    class xletojson
    {
        //Input and output folders
        public const String folderPath = @"d:\NSDM\input\";
        public const String outPath = @"d:\NSDM\output\";
        public const String CONNECTSTR = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};
                    Extended Properties=""Excel 12.0 Xml;HDR=YES"" ";

        //File extensions
        public const String INPUTFILEEXT = ".xlsx";
        public const String OUTPUTFILEEXT = ".json";

        //Conversion file names
        public const String SHELF = "Shelf.xlsx";
        public const String SCG = "SCG.xlsx";
        public const String BILLING = "Billing.xlsx";
        public const String CHANNEL = "Channel.xlsx";
        public const String SITE = "Site.xlsx";
        public const String STE = "STE.xlsx";
        public const String NI = "NI.xlsx";
        public const String LI = "LI.xlsx";
        public const String VRF = "vrf.xlsx";
        public const String LOCALCONTACT = "LocalContact.xlsx";
        public const String CUSTOMER = "Customer.xlsx";
        public const String SERVICEINSTANCE = "ServiceInstance.xlsx";
        public const String TRAIL = "Trail.xlsx";
        public const String CHILDTRAIL = "ChildTrail.xlsx";
        public const String WANPORT = "WanPort.xlsx";
        public const String LANPORT = "LanPort.xlsx";
        public const String CONSOLEPORT = "ConsolePort.xlsx";
        public const String XLSM = "cdfa.xlsx";
        public const String TMSVLAN = "tmsvlan.xlsx";
        public const String LOGICAL = "logical.xlsx";


        static void Main(string[] args)
        {
            foreach (string file in Directory.EnumerateFiles(folderPath, "*" + INPUTFILEEXT))
            {
                switch (Path.GetFileName(file))
                {
                    case SHELF:
                        Shelfjson();
                        break;
                    case BILLING:
                        Billingjson();
                        break;
                    case CHANNEL:
                        Channeljson();
                        break;
                    case SITE:
                        Sitejson();
                        break;
                    case STE:
                        STEjson();
                        break;
                    case NI:
                        NIjson();
                        break;
                    case LI:
                        LIjson();
                        break;
                    case VRF:
                        VRFjson();
                        break;
                    case SCG:
                        SCGjson();
                        break;
                    case LOCALCONTACT:
                        LocalContactjson();
                        break;
                    case CUSTOMER:
                        Customerjson();
                        break;
                    case SERVICEINSTANCE:
                        ServiceInstancejson();
                        break;
                    case TRAIL:
                        Trailjson();
                        break;
                    case CHILDTRAIL:
                        ChildTrailjson();
                        break;
                    case WANPORT:
                    case LANPORT:
                    case CONSOLEPORT:
                        Portjson(Path.GetFileName(file));
                        break;
                    case XLSM:
                        Xlsm();
                        break;
                    case TMSVLAN:
                        TmsVlanjson();
                        break;
                    case LOGICAL:
                        Logicaljson();
                        break;
                }
            }
        }
        private static void Shelfjson()
        {
             //Shelf file info
            var pathToExcel = folderPath + SHELF;
            var sheetName = "SHELF";
            int skipline = 6;
            var destinationPath = outPath + SHELF.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            var gName = "Routing/Protocol Info";
            var gName2 = "Device General Info";
            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json, json2, json3;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.ShelfObject
                        {
                            type1 = row[1].ToString().TrimEnd() == "" ? null : "oci/shelf",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            action = "create",
                            type = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            status = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
                            vendor = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            model = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            rev = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            networkId = row[8].ToString().TrimEnd() == "" ? null : row[8].ToString().TrimEnd(),
                            dimensionUnits = row[1].ToString().TrimEnd() == "" ? null : "INCHES",
                            barCode = row[27].ToString().TrimEnd() == "" ? null : row[27].ToString().TrimEnd(),
                            batchNumber = row[28].ToString().TrimEnd() == "" ? null : row[28].ToString().TrimEnd(),
                            purchaseDate = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
                            purchasePrice = row[30].ToString().TrimEnd() == "" ? null : row[30].ToString().TrimEnd(),
                            serialNumber = row[31].ToString().TrimEnd() == "" ? null : row[31].ToString().TrimEnd(),
                            orderNum = row[32].ToString().TrimEnd() == "" ? null : row[32].ToString().TrimEnd(),
                            installedDate = row[20].ToString().TrimEnd() == "" ? null : row[20].ToString().TrimEnd(),
                            comments = row[42].ToString().TrimEnd() == "" ? null : row[42].ToString().TrimEnd(),
                            targetId = row[45].ToString().TrimEnd() == "" ? null : row[45].ToString().TrimEnd(),
                            parentSite = row[47].ToString().TrimEnd() == "" ? null : new List<string> { row[47].ToString().TrimEnd() },
                            ServiceInstance = row[50].ToString().TrimEnd() == "" ? null : new List<string> { row[50].ToString().TrimEnd() },
                            dynamicAttributes = (row[85].ToString().TrimEnd() + row[91].ToString().TrimEnd() + row[94].ToString().TrimEnd() + row[79].ToString().TrimEnd() + row[95].ToString().TrimEnd()) == "" ? null : new List<nsdm.DynamicAttribute>
                            {
                                row[85].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Comments",
                                    attributeValue = row[85].ToString().TrimEnd()
                                },
                                row[91].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Loopback IP Address",
                                    attributeValue = row[91].ToString().TrimEnd()
                                },
                                row[94].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "LAN IP Address",
                                    attributeValue = row[94].ToString().TrimEnd()
                                },
                                row[92].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Other Protocol Description",
                                    attributeValue = row[92].ToString().TrimEnd()
                                },
                                 row[95].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Firewall IP Address",
                                    attributeValue = row[95].ToString().TrimEnd()
                                },
                                 row[84].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Site Redundancy Plan",
                                    attributeValue = row[4].ToString().TrimEnd()
                                },
                                 row[90].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "IP Address",
                                    attributeValue = row[90].ToString().TrimEnd()
                                },
                                 row[94].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "WAN IP Address",
                                    attributeValue = row[94].ToString().TrimEnd()
                                },
                                 row[68].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "TMS",
                                    attributeValue = row[68].ToString().TrimEnd()
                                },
                                 row[69].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Cloud",
                                    attributeValue = row[69].ToString().TrimEnd()
                                },
                                 row[82].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Routing Flags",
                                    attributeValue = row[82].ToString().TrimEnd()
                                },
                                 row[83].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "SNMP Community Strings",
                                    attributeValue = row[83].ToString().TrimEnd()
                                },
                                 row[70].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "IS DLSW/SNA?",
                                    attributeValue = row[70].ToString().TrimEnd()
                                },
                                 row[72].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "IS SNA Serial?",
                                    attributeValue = row[72].ToString().TrimEnd()
                                },
                                 row[73].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "IS SNA Ethernet?",
                                    attributeValue = row[73].ToString().TrimEnd()
                                },
                                 row[75].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "IS Encryption?",
                                    attributeValue = row[75].ToString().TrimEnd()
                                },
                                 row[93].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "IS Other Protocol?",
                                    attributeValue = "NO"
                                },
                                 row[92].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Secondary IP Address",
                                    attributeValue = row[92].ToString().TrimEnd()
                                },
                                 row[51].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "1MB Number",
                                    attributeValue = row[51].ToString().TrimEnd()
                                },
                                 row[52].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "1MB Owner",
                                    attributeValue = row[52].ToString().TrimEnd()
                                },
                                 row[58].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "Managed By",
                                    attributeValue = row[58].ToString().TrimEnd()
                                },
                                 row[65].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "Support Group",
                                    attributeValue = row[65].ToString().TrimEnd()
                                },
                                 row[59].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "Equipment Model",
                                    attributeValue = row[59].ToString().TrimEnd()
                                },
                                 row[60].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "Maintenance Class",
                                    attributeValue = row[60].ToString().TrimEnd()
                                },
                                 row[53].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "Shared Flag",
                                    attributeValue = row[53].ToString().TrimEnd()
                                }
                            }.Where(x => x != null).ToList()

                        };

                        //Generates Shelf JSON from the LINQ query
                        json = JsonConvert.SerializeObject(query.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                                {
                                    NullValueHandling = NullValueHandling.Ignore
                                });
                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }
                    foreach (JToken tk in temp.Descendants())
                    {

                        if (tk.Type == JTokenType.Property)
                        {
                            JProperty p = tk as JProperty;

                            if (p.Name == "purchaseDate" || p.Name == "installedDate")
                            {
                                try
                                {
                                    DateTime p1 = DateTime.Parse(p.Value.ToString());
                                    p.Value = p1.ToString("yyyy-MM-ddT00:00:00.000Z");
                                }
                                catch
                                {
                                    p.Value = "Error";
                                    Debug.WriteLine("Fail to convert " + p.Name);
                                }
                            }

                        }
                    }

                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[47].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[47].ToString().TrimEnd() == "" ? null : row[47].ToString().TrimEnd(),
                            name = row[47].ToString().TrimEnd() == "" ? null : row[47].ToString().TrimEnd(),
                            action1 = row[47].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }
                    using (var rdr3 = cmd.ExecuteReader())
                    {

                        var query3 =
                        from DbDataRecord row in rdr3
                        //orderby row[50] descending
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[50].ToString().TrimEnd() == "" ? null : "otech/serviceInstance",
                            refId = row[50].ToString().TrimEnd() == "" ? null : row[50].ToString().TrimEnd(),
                            name = row[50].ToString().TrimEnd() == "" ? null : row[50].ToString().TrimEnd(),
                            action1 = row[50].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json3 = JsonConvert.SerializeObject(query3.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore
                        });
                    }

                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row
                    while (temp2.Last.Count() == 0)
                    {
                        temp2.Remove(temp2.Last);
                    }

                    JArray temp3 = JArray.Parse(json3);
                    //remove blank row
                    while (temp3.FirstOrDefault() != null && temp3.Last.Count() == 0)
                    {
                        temp3.Remove(temp3.Last);
                    }

                    //foreach (var item in temp3.Children())
                    //{
                    //    if (item.Count() == 0)
                    //        temp3.Remove(temp3.item.);
                    //}


                    var test1 = new JArray((temp2.Union(temp3)).Union(temp));
                    json = test1.ToString();

                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void Channeljson()
        {
            //Shelf file info
            var pathToExcel = folderPath + CHANNEL;
            var sheetName = "Channel";
            var destinationPath = outPath + CHANNEL.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.ChannelObject
                        {
                            name = row[1].ToString(),
                            band_width = row[2].ToString(),
                            status = row[4].ToString(),

                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        //private static void Billingjson()
        //{
        //    //Shelf file info
        //    var pathToExcel = folderPath + BILLING;
        //    var sheetName = "Billing Data";
        //    var destinationPath = outPath + BILLING.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
        //    //***********************************************************************

        //    //Use this connection string if you have Office 2007+ drivers installed and 
        //    //your data is saved in a .xlsx file
        //    var connectionString = String.Format(CONNECTSTR, pathToExcel);

        //    //Creating and opening a data connection to the Excel sheet 
        //    using (var conn = new OleDbConnection(connectionString))
        //    {
        //        conn.Open();
        //        string json;
        //        var cmd = conn.CreateCommand();
        //        cmd.CommandText = String.Format(
        //            @"SELECT * FROM [{0}$]",
        //            sheetName);

        //        using (var rdr = cmd.ExecuteReader())
        //        {

        //            //LINQ query - when executed will create anonymous objects for each row
        //            var query =
        //                from DbDataRecord row in rdr
        //                select new nsdm.BillingObject
        //                {
        //                    billingCust = new nsdm.BillingCustObject
        //                    {
        //                        type1 = row[32].ToString().TrimEnd() == "" ? null : "oci/customer",
        //                        refId = row[32].ToString().TrimEnd() == "" ? null : row[32].ToString().TrimEnd(),
        //                        name = row[32].ToString().TrimEnd() == "" ? null : row[32].ToString().TrimEnd(),
        //                        action1 = row[32].ToString().TrimEnd() == "" ? null : "noUpdate"
        //                    },
        //                    billingService = new nsdm.BillingServiceObject
        //                    {
        //                        type1 = row[33].ToString().TrimEnd() == "" ? null : "otech/serviceInstance",
        //                        refId = row[33].ToString().TrimEnd() == "" ? null : row[33].ToString().TrimEnd(),
        //                        name = row[33].ToString().TrimEnd() == "" ? null : row[33].ToString().TrimEnd(),
        //                        action1 = row[33].ToString().TrimEnd() == "" ? null : "noUpdate"
        //                    },
        //                    billingData = new nsdm.BillingDataObject
        //                    {
        //                        type1 = row[35].ToString().TrimEnd() == "" ? null : "otech/BillingData",
        //                        refId = row[35].ToString().TrimEnd() == "" ? null : row[35].ToString().TrimEnd(),
        //                        action1 = row[10].ToString().TrimEnd() == "" ? null : "create",
        //                        datasource = row[24].ToString().TrimEnd() == "" ? null : row[24].ToString().TrimEnd(),
        //                        srNo = row[26].ToString().TrimEnd() == "" ? null : row[26].ToString().TrimEnd(),
        //                        compCode = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
        //                        resourceType1 = row[16].ToString().TrimEnd() == "" ? null : row[16].ToString().TrimEnd(),
        //                        shiftCode = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
        //                        periodicity = row[11].ToString().TrimEnd() == "" ? null : row[11].ToString().TrimEnd(),
        //                        sharingPer = row[17].ToString().TrimEnd() == "" ? null : row[17].ToString().TrimEnd(),
        //                        effectiveDate = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd(),
        //                        deInstallDate = row[25].ToString().TrimEnd() == "" ? null : row[25].ToString().TrimEnd(),
        //                        chargeType = row[8].ToString().TrimEnd() == "" ? null : row[8].ToString().TrimEnd(),
        //                        chargeCategory = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
        //                        billAction = row[31].ToString().TrimEnd() == "" ? null : row[31].ToString().TrimEnd(),
        //                        qty = row[12].ToString().TrimEnd() == "" ? null : row[12].ToString().TrimEnd(),
        //                        rate = row[13].ToString().TrimEnd() == "" ? null : row[13].ToString().TrimEnd(),
        //                        calnetProductId = row[9].ToString().TrimEnd() == "" ? null : row[9].ToString().TrimEnd(),
        //                        vendorAccount = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
        //                        circuitId = row[18].ToString().TrimEnd() == "" ? null : row[18].ToString().TrimEnd(),
        //                        deviceName = row[20].ToString().TrimEnd() == "" ? null : row[20].ToString().TrimEnd(),
        //                        deviceType = row[23].ToString().TrimEnd() == "" ? null : row[23].ToString().TrimEnd(),
        //                        deviceSerialNum = row[21].ToString().TrimEnd() == "" ? null : row[21].ToString().TrimEnd(),
        //                        deviceModel = row[19].ToString().TrimEnd() == "" ? null : row[19].ToString().TrimEnd(),
        //                        phoneNumber = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
        //                        recordName = row[22].ToString().TrimEnd() == "" ? null : row[22].ToString().TrimEnd(),
        //                        lastModifiedDate = row[27].ToString().TrimEnd() == "" ? null : row[27].ToString().TrimEnd(),
        //                        //lastUpdatedBy = row[28].ToString().TrimEnd() == "" ? null : row[28].ToString().TrimEnd(),
        //                        comments = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
        //                        //commentsTemp = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
        //                        //commentsHistory = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
        //                        billExpiryDate = row[30].ToString().TrimEnd() == "" ? null : row[30].ToString().TrimEnd(),
        //                        rateDisplay = row[15].ToString().TrimEnd() == "" ? null : row[15].ToString().TrimEnd(),
        //                        rawRate = row[14].ToString().TrimEnd() == "" ? null : row[14].ToString().TrimEnd(),
        //                        isRateUpdate = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
        //                        customer = row[32].ToString().TrimEnd() == "" ? null : new List<string> { (row[32].ToString().TrimEnd()) },
        //                        ServiceInstance = row[33].ToString().TrimEnd() == "" ? null : new List<string> { row[33].ToString().TrimEnd() }
                            
        //                }
        //            };
        //            //Generates Shelf JSON from the LINQ query
        //            json = JsonConvert.SerializeObject(query.Skip(5), Formatting.Indented, new JsonSerializerSettings
        //            {
        //                NullValueHandling = NullValueHandling.Ignore,
        //                DefaultValueHandling = DefaultValueHandling.Ignore
        //            });


        //            //JArray json3 = new JArray(temp.Union(unique));
        //            json = json.ToString();
        //            //Write the file to the destination path    
        //            File.WriteAllText(destinationPath, json);
        //        }
        //    }
        //}
        private static void Billingjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + BILLING;
            var sheetName = "Billing Data";
            var destinationPath = outPath + BILLING.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            int skipline = 5;
            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                string json2;
                string json3;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.BillingDataObject
                        {
                            type1 = row[36].ToString().TrimEnd() == "" ? null : "otech/BillingData",
                            refId = row[36].ToString().TrimEnd() == "" ? null : row[36].ToString().TrimEnd(),
                            action1 = row[10].ToString().TrimEnd() == "" ? null : "create",
                            datasource = row[24].ToString().TrimEnd() == "" ? null : row[24].ToString().TrimEnd(),
                            srNo = row[26].ToString().TrimEnd() == "" ? null : row[26].ToString().TrimEnd(),
                            compCode = row[9].ToString().TrimEnd() == "" ? null : row[9].ToString().TrimEnd(),
                            resourceType1 = row[16].ToString().TrimEnd() == "" ? null : row[16].ToString().TrimEnd(),
                            shiftCode = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
                            periodicity = row[11].ToString().TrimEnd() == "" ? null : row[11].ToString().TrimEnd(),
                            sharingPer = row[17].ToString().TrimEnd() == "" ? null : row[17].ToString().TrimEnd(),
                            effectiveDate = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd(),
                            deInstallDate = row[25].ToString().TrimEnd() == "" ? null : row[25].ToString().TrimEnd(),
                            chargeType = row[8].ToString().TrimEnd() == "" ? null : row[8].ToString().TrimEnd(),
                            chargeCategory = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
                            billAction = row[31].ToString().TrimEnd() == "" ? null : row[31].ToString().TrimEnd(),
                            qty = row[12].ToString().TrimEnd() == "" ? null : row[12].ToString().TrimEnd(),
                            rate = row[13].ToString().TrimEnd() == "" ? null : row[13].ToString().TrimEnd(),
                            calnetProductId = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            vendorAccount = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
                            circuitId = row[18].ToString().TrimEnd() == "" ? null : row[18].ToString().TrimEnd(),
                            deviceName = row[20].ToString().TrimEnd() == "" ? null : row[20].ToString().TrimEnd(),
                            deviceType = row[23].ToString().TrimEnd() == "" ? null : row[23].ToString().TrimEnd(),
                            deviceSerialNum = row[21].ToString().TrimEnd() == "" ? null : row[21].ToString().TrimEnd(),
                            deviceModel = row[19].ToString().TrimEnd() == "" ? null : row[19].ToString().TrimEnd(),
                            phoneNumber = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            recordName = row[22].ToString().TrimEnd() == "" ? null : row[22].ToString().TrimEnd(),
                            //lastModifiedDate = row[27].ToString().TrimEnd() == "" ? null : row[27].ToString().TrimEnd(),
                            lastUpdatedBy = row[28].ToString().TrimEnd() == "" ? null : row[28].ToString().TrimEnd(),
                            previousSrNo = row[32].ToString().TrimEnd() == "" ? null : row[32].ToString().TrimEnd(),
                            comments = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            commentsTemp = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            //commentsTemp = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            //commentsHistory = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            billExpiryDate = row[30].ToString().TrimEnd() == "" ? null : row[30].ToString().TrimEnd(),
                            rateDisplay = row[15].ToString().TrimEnd() == "" ? null : row[15].ToString().TrimEnd(),
                            //rawRate = row[14].ToString().TrimEnd() == "" ? null : row[14].ToString().TrimEnd(),
                            //isRateUpdate = row[32].ToString().TrimEnd() == "" ? null : "No",
                            customer = row[33].ToString().TrimEnd() == "" ? null : new List<string> { (row[33].ToString().TrimEnd()) },
                            ServiceInstance = row[34].ToString().TrimEnd() == "" ? null : new List<string> { row[34].ToString().TrimEnd() }
                        };
                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        DefaultValueHandling = DefaultValueHandling.Ignore
                    });
                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[33].ToString().TrimEnd() == "" ? null : "oci/customer",
                            refId = row[33].ToString().TrimEnd() == "" ? null : row[33].ToString().TrimEnd(),
                            name = row[33].ToString().TrimEnd() == "" ? null : row[33].ToString().TrimEnd(),
                            action1 = row[33].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr3 = cmd.ExecuteReader())
                    {

                        var query3 =
                        from DbDataRecord row in rdr3
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[34].ToString().TrimEnd() == "" ? null : "otech/serviceInstance",
                            refId = row[34].ToString().TrimEnd() == "" ? null : row[34].ToString().TrimEnd(),
                            name = row[34].ToString().TrimEnd() == "" ? null : row[34].ToString().TrimEnd(),
                            action1 = row[34].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json3 = JsonConvert.SerializeObject(query3.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }

                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row
                    while (temp2.Last.Count() == 0)
                    {
                        temp2.Remove(temp2.Last);
                    }

                    JArray temp3 = JArray.Parse(json3);
                    //remove blank row
                    while (temp3.Last.Count() == 0)
                    {
                        temp3.Remove(temp3.Last);
                    }


                    //remove duplicate customer objects
                    var unique2 = temp2.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp2.Count() - 1; i >= 0; i--)
                    {
                        var token = temp2[i];
                        if (!unique2.Contains(token))
                        {
                            token.Remove();
                        }
                    };

                    //remove duplicate customer objects
                    var unique3 = temp3.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp3.Count() - 1; i >= 0; i--)
                    {
                        var token = temp3[i];
                        if (!unique3.Contains(token))
                        {
                            token.Remove();
                        }
                    };


                    foreach (JToken tk in temp.Descendants())
                    {

                        if (tk.Type == JTokenType.Property)
                        {
                            JProperty p = tk as JProperty;

                            if (p.Name == "effectiveDate" || p.Name == "deInstallDate" || p.Name == "billExpiryDate" || p.Name == "lastModifiedDate")
                            {
                                try
                                {
                                    //double d = double.Parse(p.Value.ToString());
                                    //DateTime p1 = DateTime.FromOADate(d);

                                    DateTime p1 = DateTime.Parse(p.Value.ToString());
                                    p.Value = p1.ToString("yyyy-MM-ddT00:00:00.000Z");
                                }
                                catch
                                {
                                    p.Value = "Error";
                                    Debug.WriteLine("Fail to convert " + p.Name);
                                }
                            }

                        }
                    }

                    var test1 = new JArray((temp2.Union(temp3)).Union(temp));
                    //var arrayOfObjects = JsonConvert.SerializeObject(
                    //     new[] { temp2, temp3, temp });

                    //JArray json3 = new JArray(temp.Union(unique));
                    json = test1.ToString();
                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void Portjson(string filename)
        {
            //Shelf file info
            var pathToExcel = folderPath + filename;
            var sheetName = "WAN PORT";
            var destinationPath = outPath + filename.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            int skipline = 6;
            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json, json2, json3;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.PortObject
                        {
                            type1 = row[1].ToString().TrimEnd() == "" ? null : "oci/port",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            bandWidth = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
                            //direction = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            //portAccessId = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            connectorType = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            //channelization = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
                            //parentPortChanName = row[8].ToString().TrimEnd() == "" ? null : row[8].ToString().TrimEnd(),
                            //networkId = row[9].ToString().TrimEnd() == "" ? null : row[9].ToString().TrimEnd(),
                            description = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd(),
                            role = row[11].ToString().TrimEnd() == "" ? null : row[11].ToString().TrimEnd(),
                            //hecig = row[13].ToString().TrimEnd() == "" ? null : row[13].ToString().TrimEnd(),
                            //networkDirection = row[14].ToString().TrimEnd() == "" ? null : row[14].ToString().TrimEnd(),
                            //physicalPortName = row[15].ToString().TrimEnd() == "" ? null : row[15].ToString().TrimEnd(),
                            //aidFormula = row[12].ToString().TrimEnd() == "" ? null : row[12].ToString().TrimEnd(),
                            status = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            //site = row[17].ToString().TrimEnd() == "" ? null : row[17].ToString().TrimEnd(),
                            //portClass = row[16].ToString().TrimEnd() == "" ? null : row[16].ToString().TrimEnd(),
                            parentShelf = row[18].ToString().TrimEnd() == "" ? null : new List<string> { row[18].ToString().TrimEnd() },
                        };
                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        DefaultValueHandling = DefaultValueHandling.Ignore
                    });
                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.PortTypeObject
                        {
                            type1 = row[17].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[17].ToString().TrimEnd() == "" ? null : row[17].ToString().TrimEnd(),
                            name = row[17].ToString().TrimEnd() == "" ? null : row[17].ToString().TrimEnd(),
                            action1 = row[17].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }
                    using (var rdr3 = cmd.ExecuteReader())
                    {

                        var query3 =
                        from DbDataRecord row in rdr3
                        select new nsdm.PortTypeObject
                        {
                            type1 = row[18].ToString().TrimEnd() == "" ? null : "oci/shelf",
                            refId = row[18].ToString().TrimEnd() == "" ? null : row[18].ToString().TrimEnd(),
                            parentSite = row[17].ToString().TrimEnd() == "" ? null : new List<string> { row[17].ToString().TrimEnd() },
                            name = row[18].ToString().TrimEnd() == "" ? null : row[18].ToString().TrimEnd(),
                            action1 = row[18].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json3 = JsonConvert.SerializeObject(query3.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }
                    //using (var rdr4 = cmd.ExecuteReader())
                    //{

                    //    var query4 =
                    //    from DbDataRecord row in rdr4
                    //    select new nsdm.PortTypeObject
                    //    {
                    //        type1 = row[16].ToString().TrimEnd() == "" ? null : "oci/portClass",
                    //        refId = row[16].ToString().TrimEnd() == "" ? null : row[16].ToString().TrimEnd(),
                    //        name = row[16].ToString().TrimEnd() == "" ? null : row[16].ToString().TrimEnd(),
                    //        action1 = row[16].ToString().TrimEnd() == "" ? null : "noUpdate"
                    //    };
                    //    //Generates Shelf JSON from the LINQ query
                    //    json4 = JsonConvert.SerializeObject(query4.Skip(6), Formatting.Indented, new JsonSerializerSettings
                    //    {
                    //        NullValueHandling = NullValueHandling.Ignore,
                    //        DefaultValueHandling = DefaultValueHandling.Ignore
                    //    });
                    //}


                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }

                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row
                    while (temp2.Last.Count() == 0)
                    {
                        temp2.Remove(temp2.Last);
                    }

                    JArray temp3 = JArray.Parse(json3);
                    //remove blank row
                    while (temp3.Last.Count() == 0)
                    {
                        temp3.Remove(temp3.Last);
                    }

                    //JArray temp4 = JArray.Parse(json4);
                    ////remove blank row
                    //while (temp4.Last.Count() == 0)
                    //{
                    //    temp4.Remove(temp4.Last);
                    //}

                    ////remove duplicate site objects
                    //var unique = temp2.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    //// Iterate backwards over the JObject to remove any duplicate keys
                    //for (int i = temp2.Count() - 1; i >= 0; i--)
                    //{
                    //    var token = temp2[i];
                    //    if (!unique.Contains(token))
                    //    {
                    //        token.Remove();
                    //    }
                    //};

                    //remove duplicate shelf objects
                    var unique = temp3.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp3.Count() - 1; i >= 0; i--)
                    {
                        var token = temp3[i];
                        if (!unique.Contains(token))
                        {
                            token.Remove();
                        }
                    };

                    ////remove duplicate portclass objects
                    //unique = temp4.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    //// Iterate backwards over the JObject to remove any duplicate keys
                    //for (int i = temp4.Count() - 1; i >= 0; i--)
                    //{
                    //    var token = temp4[i];
                    //    if (!unique.Contains(token))
                    //    {
                    //        token.Remove();
                    //    }
                    //};


                    //var test1 = new JArray(((temp2.Union(temp3)).Union(temp4)).Union(temp));
                    var test1 = new JArray((temp2.Union(temp3)).Union(temp));
                    //var arrayOfObjects = JsonConvert.SerializeObject(
                    //     new[] { temp2, temp3, temp });

                    //JArray json3 = new JArray(temp.Union(unique));
                    json = test1.ToString();
                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void Sitejson()
        {
            //Shelf file info
            var pathToExcel = folderPath + SITE;
            var sheetName = "Site";
            var destinationPath = outPath + SITE.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            var gName = "Site_General_Info";
            int skipline = 6;
            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    "SITE");

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        where (row[1] != null && row[1].ToString().TrimEnd() != "")
                        select new nsdm.SiteObject
                        {
                            type1 = row[1].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            siteID = row[1].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            type = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            status = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
                            clli = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            address = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd(),
                            city = row[11].ToString().TrimEnd() == "" ? null : row[11].ToString().TrimEnd(),
                            stateProv = row[12].ToString().TrimEnd() == "" ? null : row[12].ToString().TrimEnd(),
                            postalCode1 = row[13].ToString().TrimEnd() == "" ? null : row[13].ToString().TrimEnd(),
                            //floor = row[18].ToString().TrimEnd() == "" ? null : row[18].ToString().TrimEnd(),
                            room = row[19].ToString().TrimEnd() == "" ? null : row[19].ToString().TrimEnd(),
                            //inServiceDate = row[25].ToString().TrimEnd() == "" ? null : DateTime.Parse(row[25].ToString()).ToString("yyyyMMdd"),
                            //comments = row[42].ToString().TrimEnd() == "" ? null : row[42].ToString().TrimEnd(),
                            //parentSiteName = row[29].ToString().TrimEnd() == "" ? null : new List<string> { row[29].ToString().TrimEnd() },
                            //customer = row[30].ToString().TrimEnd() == "" ? null : new List<string> { row[30].ToString().TrimEnd() },
                            dynamicAttributes = row[6].ToString().TrimEnd() == "" ? null : new List<nsdm.DynamicAttribute>
                            {
                                row[6].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = "Primary",
                                    attributeName = "LATA",
                                    attributeValue = row[6].ToString().TrimEnd()
                                },
                                new nsdm.DynamicAttribute {
                                    groupName = "Primary",
                                    attributeName = "Critical Site",
                                    attributeValue = row[9].ToString().TrimEnd()
                                }
                            }.Where(x => x != null).ToList()
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                    {
                        //NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void STEjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + STE;
            var sheetName = "STE";
            var destinationPath = outPath + STE.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json, json2, json3, json4;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.STEObject
                        {
                            type1 = row[2].ToString().TrimEnd() == "" ? null : "otech/serviceInstance",
                            refId = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            action1 = row[4].ToString().TrimEnd() == "" ? null : "otech/processAssociations",
                            name = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            site = row[26].ToString().TrimEnd() == "" ? null : new List<string> { row[26].ToString().TrimEnd() },
                            bandwidth = row[29].ToString().TrimEnd() == "" ? null : new List<string> { row[29].ToString().TrimEnd() },
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(6), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });

                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[26].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[26].ToString().TrimEnd() == "" ? null : row[26].ToString().TrimEnd(),
                            name = row[26].ToString().TrimEnd() == "" ? null : row[26].ToString().TrimEnd(),
                            action1 = row[26].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(6), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr3 = cmd.ExecuteReader())
                    {

                        var query3 =
                        from DbDataRecord row in rdr3
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[29].ToString().TrimEnd() == "" ? null : "occ/bandwidth",
                            refId = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
                            name = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
                            action1 = row[29].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json3 = JsonConvert.SerializeObject(query3.Skip(6), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr4 = cmd.ExecuteReader())
                    {

                        var query4 =
                        from DbDataRecord row in rdr4
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[25].ToString().TrimEnd() == "" ? null : "otech/serviceInstance",
                            refId = row[25].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            name = row[25].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            action1 = row[25].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json4 = JsonConvert.SerializeObject(query4.Skip(6), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }
                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }

                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row
                    while (temp2.Last.Count() == 0)
                    {
                        temp2.Remove(temp2.Last);
                    }

                    JArray temp3 = JArray.Parse(json3);
                    //remove blank row
                    while (temp3.Last.Count() == 0)
                    {
                        temp3.Remove(temp3.Last);
                    }

                    JArray temp4 = JArray.Parse(json4);
                    //remove blank row
                    while (temp4.Last.Count() == 0)
                    {
                        temp4.Remove(temp4.Last);
                    }

                    //remove duplicate site objects
                    var unique2 = temp2.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp2.Count() - 1; i >= 0; i--)
                    {
                        var token = temp2[i];
                        if (!unique2.Contains(token))
                        {
                            token.Remove();
                        }
                    };

                    //remove duplicate bandwidth objects
                    var unique3 = temp3.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp3.Count() - 1; i >= 0; i--)
                    {
                        var token = temp3[i];
                        if (!unique3.Contains(token))
                        {
                            token.Remove();
                        }
                    };


                    //remove duplicate customer objects
                    var unique4 = temp4.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp4.Count() - 1; i >= 0; i--)
                    {
                        var token = temp4[i];
                        if (!unique4.Contains(token))
                        {
                            token.Remove();
                        }
                    };


                    //temp.Concat(temp2);
                    var test1 = new JArray(((temp2.Union(temp3)).Union(temp4)).Union(temp));
                    //var arrayOfObjects = JsonConvert.SerializeObject(
                    //     new[] { temp2, temp3, temp });

                    //JArray json3 = new JArray(temp.Union(unique));
                    json = test1.ToString();
                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void NIjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + NI;
            var destinationPath = outPath + NI.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};
                    Extended Properties=""Excel 12.0 Xml;HDR=YES""
                ", pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    "WAN-NI");

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.NIObject
                        {
                            ni_type = "WAN-NI",
                            name = row[1].ToString(),
                            type = row[2].ToString(),
                            status = row[3].ToString(),
                            parent_site_name = row[51].ToString(),
                            parent_shelf_name = row[52].ToString(),
                            port_name = row[56].ToString()
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    "LAN-NI");

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.NIObject
                        {
                            ni_type = "LAN-NI",
                            name = row[1].ToString(),
                            type = row[2].ToString(),
                            status = row[3].ToString(),
                            parent_site_name = row[51].ToString(),
                            parent_shelf_name = row[52].ToString(),
                            port_name = row[56].ToString()
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.AppendAllText(destinationPath, json);
                }
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    "Console-NI");

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.NIObject
                        {
                            ni_type = "Console-NI",
                            name = row[1].ToString(),
                            type = row[2].ToString(),
                            status = row[3].ToString(),
                            parent_site_name = row[51].ToString(),
                            parent_shelf_name = row[52].ToString(),
                            port_name = row[56].ToString()
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.AppendAllText(destinationPath, json);
                }
            }
        }
        private static void LIjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + LI;
            var sheetName = "Logical Interface";
            var destinationPath = outPath + LI.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.LIObject
                        {
                            name = row[1].ToString(),
                            type = row[2].ToString(),
                            status = row[3].ToString(),
                            installed_date = row[65].ToString().TrimEnd() == "" ? null : DateTime.Parse(row[65].ToString()).ToString("yyyyMMdd"),

                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void VRFjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + VRF;
            var sheetName = "VRF";
            var destinationPath = outPath + VRF.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.VRFObject
                        {
                            type1 = row[1].ToString().TrimEnd() == "" ? null : "oct/vrf",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            type = row[1].ToString().TrimEnd() == "" ? null : "Full Mesh",
                            status = row[1].ToString().TrimEnd() == "" ? null : "ACTIVE",
                            routeMapOut = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
                            comments = row[15].ToString().TrimEnd() == "" ? null : row[15].ToString().TrimEnd()
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(6), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void SCGjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + SCG;
            var sheetName = "Service Connection Group";
            var destinationPath = outPath + SCG.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.SCGObject
                        {
                            name = row[1].ToString(),
                            type = row[2].ToString(),
                            status = row[3].ToString(),

                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void LocalContactjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + LOCALCONTACT;
            var sheetName = "Local Contact";
            var destinationPath = outPath + LOCALCONTACT.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json, json2;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.LocalContactObject
                        {
                            type1 = "nsdm/LocalContact",
                            refId = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            name = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            type = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            description = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
                            phoneNumber = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            alternatePhoneNumber = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            emailAddress = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            site = row[7].ToString().TrimEnd() == "" ? null : new List<string> { row[7].ToString().TrimEnd() },

                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(5), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });

                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[7].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
                            name = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
                            action1 = row[7].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(5), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }

                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row
                    while (temp2.Last.Count() == 0)
                    {
                        temp2.Remove(temp2.Last);
                    }

                    //remove duplicate customer objects
                    var unique2 = temp2.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp2.Count() - 1; i >= 0; i--)
                    {
                        var token = temp2[i];
                        if (!unique2.Contains(token))
                        {
                            token.Remove();
                        }
                    };

                    var test1 = new JArray(temp2.Union(temp));
                    //var arrayOfObjects = JsonConvert.SerializeObject(
                    //     new[] { temp2, temp3, temp });

                    //JArray json3 = new JArray(temp.Union(unique));
                    json = test1.ToString();

                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void Customerjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + CUSTOMER;
            var sheetName = "Customer";
            var destinationPath = outPath + CUSTOMER.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.CustomerObject
                        {
                            type1 = row[1].ToString().TrimEnd() == "" ? null : "oci/customer",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            customerName = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            status = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
                            type = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            //primaryPhoneNo = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            //billingAddress = row[14].ToString().TrimEnd() == "" ? null : row[14].ToString().TrimEnd(),
                            //city = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            //stateProv = row[13].ToString().TrimEnd() == "" ? null : row[13].ToString().TrimEnd(),
                            //postalCode1 = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd(),
                            //postalCode2 = row[11].ToString().TrimEnd() == "" ? null : row[11].ToString().TrimEnd(),
                            //country = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
                            //npaNxx = row[9].ToString().TrimEnd() == "" ? null : row[9].ToString().TrimEnd(),
                            //floor = row[8].ToString().TrimEnd() == "" ? null : row[8].ToString().TrimEnd(),
                            //room = row[12].ToString().TrimEnd() == "" ? null : row[12].ToString().TrimEnd(),
                            //billingCode = row[15].ToString().TrimEnd() == "" ? null : row[15].ToString().TrimEnd(),
                            //billingFax = row[16].ToString().TrimEnd() == "" ? null : row[16].ToString().TrimEnd(),
                            //comments = row[17].ToString().TrimEnd() == "" ? null : row[17].ToString().TrimEnd(),
                            //billingContacts = row[18].ToString().TrimEnd() == "" ? null : row[18].ToString().TrimEnd(),
                            //techContacts = row[19].ToString().TrimEnd() == "" ? null : row[19].ToString().TrimEnd()
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(3), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });

                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }
                    json = temp.ToString();
                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void ServiceInstancejson()
        {
            //Shelf file info
            var pathToExcel = folderPath + SERVICEINSTANCE;
            var sheetName = "SI";
            var destinationPath = outPath + SERVICEINSTANCE.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            int skipline = 4;

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json, json2, json2b, json3, json4, json5;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.ServiceInstanceObject
                        {
                            type1 = row[2].ToString().TrimEnd() == "" ? null : "otech/serviceInstance",
                            refId = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            name = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            status = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            isparentServ = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            serviceType = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            serviceStatus = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            description = row[8].ToString().TrimEnd() == "" ? null : row[8].ToString().TrimEnd(),
                            svcRequestNumber = row[7].ToString().TrimEnd() == "" ? null : row[7].ToString().TrimEnd(),
                            comments = row[9].ToString().TrimEnd() == "" ? null : row[9].ToString().TrimEnd(),
                            vendor = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd(),
                            CustAccountCode = row[11].ToString().TrimEnd() == "" ? null : row[11].ToString().TrimEnd(),
                            parentServiceName = row[12].ToString().TrimEnd() == "" ? null : row[12].ToString().TrimEnd(),
                            installedDate = row[13].ToString().TrimEnd() == "" ? null : row[13].ToString().TrimEnd(),
                            decommissionDate = row[16].ToString().TrimEnd() == "" ? null : row[16].ToString().TrimEnd(),
                            lanActivationDate = row[18].ToString().TrimEnd() == "" ? null : row[18].ToString().TrimEnd(),
                            srCompletionDate = row[20].ToString().TrimEnd() == "" ? null : row[20].ToString().TrimEnd(),
                            srDecommissionDate = row[21].ToString().TrimEnd() == "" ? null : row[21].ToString().TrimEnd(),
                            site = row[26].ToString().TrimEnd() == "" ? null : new List<string> { row[26].ToString().TrimEnd() },
                            customer = row[25].ToString().TrimEnd() == "" ? null : new List<string> { row[25].ToString().TrimEnd() },
                            trail = row[28].ToString().TrimEnd() == "" ? null : new List<string> { row[28].ToString().TrimEnd() },
                            bandwidth = row[29].ToString().TrimEnd() == "" ? null : new List<string> { row[29].ToString().TrimEnd() },
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });

                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[26].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[26].ToString().TrimEnd() == "" ? null : row[26].ToString().TrimEnd(),
                            name = row[26].ToString().TrimEnd() == "" ? null : row[26].ToString().TrimEnd(),
                            action1 = row[26].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr2b = cmd.ExecuteReader())
                    {

                        var query2b =
                        from DbDataRecord row in rdr2b
                        select new nsdm.SIShelfObject
                        {
                            type1 = row[27].ToString().TrimEnd() == "" ? null : "oci/shelf",
                            refId = row[27].ToString().TrimEnd() == "" ? null : row[27].ToString().TrimEnd(),
                            parentSite = row[27].ToString().TrimEnd() == "" ? null : new List<string> { row[27].ToString().TrimEnd() },
                            action1 = row[27].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2b = JsonConvert.SerializeObject(query2b.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr3 = cmd.ExecuteReader())
                    {

                        var query3 =
                        from DbDataRecord row in rdr3
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[29].ToString().TrimEnd() == "" ? null : "occ/bandwidth",
                            refId = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
                            name = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
                            action1 = row[29].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json3 = JsonConvert.SerializeObject(query3.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr4 = cmd.ExecuteReader())
                    {

                        var query4 =
                        from DbDataRecord row in rdr4
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[25].ToString().TrimEnd() == "" ? null : "oci/customer",
                            refId = row[25].ToString().TrimEnd() == "" ? null : row[25].ToString().TrimEnd(),
                            name = row[25].ToString().TrimEnd() == "" ? null : row[25].ToString().TrimEnd(),
                            action1 = row[25].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json4 = JsonConvert.SerializeObject(query4.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr5 = cmd.ExecuteReader())
                    {

                        var query5 =
                        from DbDataRecord row in rdr5
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[28].ToString().TrimEnd() == "" ? null : "occ/trail",
                            refId = row[28].ToString().TrimEnd() == "" ? null : row[28].ToString().TrimEnd(),
                            name = row[28].ToString().TrimEnd() == "" ? null : row[28].ToString().TrimEnd(),
                            action1 = row[28].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json5 = JsonConvert.SerializeObject(query5.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }

                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row
                    while (temp2.Last.Count() == 0)
                    {
                        temp2.Remove(temp2.Last);
                    }

                    JArray temp2b = JArray.Parse(json2b);
                    //remove blank row
                    while (temp2b.Last.Count() == 0)
                    {
                        temp2b.Remove(temp2b.Last);
                    }

                    JArray temp3 = JArray.Parse(json3);
                    //remove blank row
                    while (temp3.Last.Count() == 0)
                    {
                        temp3.Remove(temp3.Last);
                    }

                    JArray temp4 = JArray.Parse(json4);
                    //remove blank row
                    while (temp4.Last.Count() == 0)
                    {
                        temp4.Remove(temp4.Last);
                    }

                    JArray temp5 = JArray.Parse(json5);
                    //remove blank row
                    while (temp5.Last.Count() == 0)
                    {
                        temp5.Remove(temp5.Last);
                    }

                    //remove duplicate site objects
                    var unique2 = temp2.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp2.Count() - 1; i >= 0; i--)
                    {
                        var token = temp2[i];
                        if (!unique2.Contains(token))
                        {
                            token.Remove();
                        }
                    };


                    //remove duplicate site objects
                    var unique2b = temp2b.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp2b.Count() - 1; i >= 0; i--)
                    {
                        var token = temp2b[i];
                        if (!unique2b.Contains(token))
                        {
                            token.Remove();
                        }
                    };

                    //remove duplicate bandwidth objects
                    var unique3 = temp3.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp3.Count() - 1; i >= 0; i--)
                    {
                        var token = temp3[i];
                        if (!unique3.Contains(token))
                        {
                            token.Remove();
                        }
                    };


                    //remove duplicate customer objects
                    var unique4 = temp4.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp4.Count() - 1; i >= 0; i--)
                    {
                        var token = temp4[i];
                        if (!unique4.Contains(token))
                        {
                            token.Remove();
                        }
                    };

                    foreach (JToken tk in temp.Descendants())
                    {

                        if (tk.Type == JTokenType.Property)
                        {
                            JProperty p = tk as JProperty;

                            if (p.Name == "decommissionDate" || p.Name == "installedDate" || p.Name == "lanActivationDate" || p.Name == "srCompletionDate" || p.Name == "srDecommissionDate")
                            {
                                try
                                {
                                    //double d = double.Parse(p.Value.ToString());
                                    //DateTime p1 = DateTime.FromOADate(d);
                                    DateTime p1 = DateTime.Parse(p.Value.ToString());
                                    p.Value = p1.ToString("yyyy-MM-ddT00:00:00.000Z");
                                }
                                catch
                                {
                                    p.Value = "Error";
                                    Debug.WriteLine("Fail to convert " + p.Name);
                                }
                            }

                        }
                    }

                    //temp.Concat(temp2);
                    var test1 = new JArray(((((temp2.Union(temp2b)).Union(temp3)).Union(temp4)).Union(temp5)).Union(temp));
                    //var arrayOfObjects = JsonConvert.SerializeObject(
                    //     new[] { temp2, temp3, temp });

                    //JArray json3 = new JArray(temp.Union(unique));
                    json = test1.ToString();
                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void Trailjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + TRAIL;
            var sheetName = "TRAIL";
            var destinationPath = outPath + TRAIL.Replace(INPUTFILEEXT, OUTPUTFILEEXT);
            var gName = "Trail General Info";
            var gName2 = "Wholesale Circuit";
            int skipline = 6;

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json, json2, json3, json4;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.TrailObject
                        {
                            type2 = row[1].ToString().TrimEnd() == "" ? null : "occ/trail",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            action1 = row[1].ToString().TrimEnd() == "" ? null : "create",
                            type = row[2].ToString().TrimEnd() == "" ? null : row[2].ToString().TrimEnd(),
                            status = row[3].ToString().TrimEnd() == "" ? null : row[3].ToString().TrimEnd(),
                            assignmentType = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            bandWidth = row[4].ToString().TrimEnd() == "" ? null : row[4].ToString().TrimEnd(),
                            protectionType = row[6].ToString().TrimEnd() == "" ? null : row[6].ToString().TrimEnd(),
                            owner = row[13].ToString().TrimEnd() == "" ? null : row[13].ToString().TrimEnd(),
                            direction = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd(),
                            billingCode = row[24].ToString().TrimEnd() == "" ? null : row[24].ToString().TrimEnd(),
                            comments = row[28].ToString().TrimEnd() == "" ? null : row[28].ToString().TrimEnd(),
                            decommissionDate = row[29].ToString().TrimEnd() == "" ? null : row[29].ToString().TrimEnd(),
                            installedDate = row[32].ToString().TrimEnd() == "" ? null : row[32].ToString().TrimEnd(),
                            zSideSite = row[36].ToString().TrimEnd() == "" ? null : new List<string> { row[36].ToString().TrimEnd() },
                            aSideSite = row[35].ToString().TrimEnd() == "" ? null : new List<string> { row[35].ToString().TrimEnd() },
                            ServiceInstance = row[38].ToString().TrimEnd() == "" ? null : new List<string> { row[38].ToString().TrimEnd() },
                            dynamicAttributes = (row[39].ToString().TrimEnd() + row[41].ToString().TrimEnd() + row[42].ToString().TrimEnd() + row[45].ToString().TrimEnd() + row[46].ToString().TrimEnd() 
                            + row[47].ToString().TrimEnd() + row[48].ToString().TrimEnd() + row[49].ToString().TrimEnd() + row[50].ToString().TrimEnd() + row[43].ToString().TrimEnd() + row[51].ToString().TrimEnd()
                            + row[52].ToString().TrimEnd() + row[55].ToString().TrimEnd() + row[44].ToString().TrimEnd() + row[56].ToString().TrimEnd()) == "" ? null : new List<nsdm.DynamicAttribute>
                            {
                                row[39].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Telco Service Code",
                                    attributeValue = row[39].ToString().TrimEnd()
                                },
                                row[41].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Media Type",
                                    attributeValue = row[41].ToString().TrimEnd()
                                },
                                row[42].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Installed by",
                                    attributeValue = row[42].ToString().TrimEnd()
                                },
                                row[45].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "CGEN Service ID",
                                    attributeValue = row[45].ToString().TrimEnd()
                                },
                                 row[46].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "USOC Code",
                                    attributeValue = row[46].ToString().TrimEnd()
                                },
                                 row[47].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Product Code",
                                    attributeValue = row[47].ToString().TrimEnd()
                                },
                                 row[48].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Vendor Account",
                                    attributeValue = row[48].ToString().TrimEnd()
                                },
                                 row[49].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Speed Requested",
                                    attributeValue = row[49].ToString().TrimEnd()
                                },
                                 row[50].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Speed Configured",
                                    attributeValue = row[50].ToString().TrimEnd()
                                },
                                 row[43].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Provider",
                                    attributeValue = row[43].ToString().TrimEnd()
                                },
                                 row[51].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Record Created Date",
                                    attributeValue = row[51].ToString().TrimEnd()
                                },
                                 row[52].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Record Created User",
                                    attributeValue = row[52].ToString().TrimEnd()
                                },
                                 row[55].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Primary Circuit Flag",
                                    attributeValue = row[55].ToString().TrimEnd()
                                },
                                 row[44].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "Vendor Feature ID",
                                    attributeValue = row[44].ToString().TrimEnd()
                                },
                                 row[56].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName,
                                    attributeName = "TSP information",
                                    attributeValue = row[56].ToString().TrimEnd()
                                },
                                 row[57].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "Wholesale Vendor",
                                    attributeValue = row[57].ToString().TrimEnd()
                                },
                                 row[58].ToString().TrimEnd() == "" ? null : new nsdm.DynamicAttribute
                                {
                                    groupName = gName2,
                                    attributeName = "Wholesale Circuit ID",
                                    attributeValue = row[58].ToString().TrimEnd()
                                }
                            }.Where(x => x != null).ToList()

                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        DefaultValueHandling = DefaultValueHandling.Ignore
                    });
                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }
                    foreach (JToken tk in temp.Descendants())
                    {

                        if (tk.Type == JTokenType.Property)
                        {
                            JProperty p = tk as JProperty;

                            if (p.Name == "decommissionDate"  || p.Name == "installedDate")
                            {
                                try
                                {
                                    DateTime p1 = DateTime.Parse(p.Value.ToString());
                                    p.Value = p1.ToString("yyyy-MM-ddT00:00:00.000Z");
                                }
                                catch
                                {
                                    p.Value = "Error";
                                    Debug.WriteLine("Fail to convert " + p.Name);
                                }
                            }
                            //if (p.Name == "installedDate")
                            //{
                            //    try
                            //    {
                            //        string sYear = p.Value.ToString().Substring(0, 4);
                            //        string sMonth = p.Value.ToString().Substring(4, 2);
                            //        string sDay = p.Value.ToString().Substring(6, 2);
                            //        p.Value = sYear + "-" + sMonth + "-" + sDay + "T00:00:00.000Z";
                            //    }
                            //    catch
                            //    {
                            //        p.Value = "Error";
                            //        Debug.WriteLine("Fail to convert " + p.Name);
                            //    }
                            //}

                        }
                    }

                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[35].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[35].ToString().TrimEnd() == "" ? null : row[35].ToString().TrimEnd(),
                            name = row[35].ToString().TrimEnd() == "" ? null : row[35].ToString().TrimEnd(),
                            action1 = row[35].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }
                    using (var rdr3 = cmd.ExecuteReader())
                    {

                        var query3 =
                        from DbDataRecord row in rdr3
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[36].ToString().TrimEnd() == "" ? null : "oci/site",
                            refId = row[36].ToString().TrimEnd() == "" ? null : row[36].ToString().TrimEnd(),
                            name = row[36].ToString().TrimEnd() == "" ? null : row[36].ToString().TrimEnd(),
                            action1 = row[36].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json3 = JsonConvert.SerializeObject(query3.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }
                    using (var rdr4 = cmd.ExecuteReader())
                    {

                        var query4 =
                        from DbDataRecord row in rdr4
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[38].ToString().TrimEnd() == "" ? null : "otech/serviceInstance",
                            refId = row[38].ToString().TrimEnd() == "" ? null : row[38].ToString().TrimEnd(),
                            name = row[38].ToString().TrimEnd() == "" ? null : row[38].ToString().TrimEnd(),
                            action1 = row[38].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json4 = JsonConvert.SerializeObject(query4.Skip(skipline), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }
                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row

                        while (temp2.Last != null && temp2.Last.Count() == 0)
                        {
                            temp2.Remove(temp2.Last);
                        }


                    JArray temp3 = JArray.Parse(json3);
                    //remove blank row
                    while (temp3.Last != null && temp3.Last.Count() == 0)
                    {
                        temp3.Remove(temp3.Last);
                    }

                    JArray temp4 = JArray.Parse(json4);
                    //remove blank row
                    while (temp4.Last != null && temp4.Last.Count() == 0)
                    {
                        temp4.Remove(temp4.Last);
                    }

                    if (temp2 != null)
                    {
                        //remove duplicate customer objects
                        var unique2 = temp2.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                        // Iterate backwards over the JObject to remove any duplicate keys
                        for (int i = temp2.Count() - 1; i >= 0; i--)
                        {
                            var token = temp2[i];
                            if (!unique2.Contains(token))
                            {
                                token.Remove();
                            }
                        };
                    }

                    if (temp3 != null)
                    {
                        //remove duplicate customer objects
                        var unique3 = temp3.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                        // Iterate backwards over the JObject to remove any duplicate keys
                        for (int i = temp3.Count() - 1; i >= 0; i--)
                        {
                            var token = temp3[i];
                            if (!unique3.Contains(token))
                            {
                                token.Remove();
                            }
                        };
                    }

                    if (temp4 != null)
                    {
                        //remove duplicate customer objects
                        var unique4 = temp4.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                        // Iterate backwards over the JObject to remove any duplicate keys
                        for (int i = temp4.Count() - 1; i >= 0; i--)
                        {
                            var token = temp4[i];
                            if (!unique4.Contains(token))
                            {
                                token.Remove();
                            }
                        };
                    }
                    var test1 = new JArray(((temp2.Union(temp3)).Union(temp4)).Union(temp));
                    json = test1.ToString();



                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
        private static void ChildTrailjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + CHILDTRAIL;
            var sheetName = "Trail Elements";
            var destinationPath = outPath + CHILDTRAIL.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);                                          

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.ChildTrailObject
                        {
                            type1 = "oci/shelf",

                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }

        private static void Xlsm()
        {
            //Shelf file info
            var pathToExcel = folderPath + XLSM;
            var sheetName = "GovOps_v2";
            var destinationPath = outPath + XLSM.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.ChildTrailObject
                        {
                            type1 = "oci/shelf",

                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });


                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }

        private static void TmsVlanjson()
        {
            //Shelf file info
            var pathToExcel = folderPath + TMSVLAN;
            var sheetName = "TMS VLAN";
            var destinationPath = outPath + TMSVLAN.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.TmsVlanObject
                        {
                            type1 = row[1].ToString().TrimEnd() == "" ? null : "oct/vlan",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            type = row[1].ToString().TrimEnd() == "" ? null : "DEFAULT",
                            status = row[1].ToString().TrimEnd() == "" ? null : "ACTIVE",
                            comments = row[10].ToString().TrimEnd() == "" ? null : row[10].ToString().TrimEnd()
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(4), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        DefaultValueHandling = DefaultValueHandling.Ignore
                    });
                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }

                    json = temp.ToString();



                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }

        private static void Logicaljson()
        {
            //Shelf file info
            var pathToExcel = folderPath + LOGICAL;
            var sheetName = "Logical Interface";
            var destinationPath = outPath + LOGICAL.Replace(INPUTFILEEXT, OUTPUTFILEEXT);

            //***********************************************************************

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(CONNECTSTR, pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                string json, json2, json3;
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName);

                using (var rdr = cmd.ExecuteReader())
                {

                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        from DbDataRecord row in rdr
                        select new nsdm.LogicalInterfaceObject
                        {
                            type1 = row[1].ToString().TrimEnd() == "" ? null : "oct/logicalInterface",
                            refId = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            name = row[1].ToString().TrimEnd() == "" ? null : row[1].ToString().TrimEnd(),
                            type = row[1].ToString().TrimEnd() == "" ? null : "MPLS Flow Point",
                            status = row[1].ToString().TrimEnd() == "" ? null : "ACTIVE",
                            ceVlanidEvcMap = row[5].ToString().TrimEnd() == "" ? null : row[5].ToString().TrimEnd(),
                            localAsNumber = row[20].ToString().TrimEnd() == "" ? null : row[20].ToString().TrimEnd(),
                            routerIdIpAddress = row[21].ToString().TrimEnd() == "" ? null : row[21].ToString().TrimEnd(),
                            vrfs = row[24].ToString().TrimEnd() == "" ? null : new List<string> { (row[24].ToString().TrimEnd()) },
                            networkInterface = row[23].ToString().TrimEnd() == "" ? null : new List<string> { (row[23].ToString().TrimEnd()) }
                        };

                    //Generates Shelf JSON from the LINQ query
                    json = JsonConvert.SerializeObject(query.Skip(6), Formatting.Indented, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    });

                    rdr.Close();
                    using (var rdr2 = cmd.ExecuteReader())
                    {

                        var query2 =
                        from DbDataRecord row in rdr2
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[24].ToString().TrimEnd() == "" ? null : "oct/vrf",
                            refId = row[24].ToString().TrimEnd() == "" ? null : row[24].ToString().TrimEnd(),
                            name = row[24].ToString().TrimEnd() == "" ? null : row[24].ToString().TrimEnd(),
                            action1 = row[24].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json2 = JsonConvert.SerializeObject(query2.Skip(4), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    using (var rdr3 = cmd.ExecuteReader())
                    {

                        var query3 =
                        from DbDataRecord row in rdr3
                        select new nsdm.BillingCustObject
                        {
                            type1 = row[23].ToString().TrimEnd() == "" ? null : "oct/networkInterface",
                            refId = row[23].ToString().TrimEnd() == "" ? null : row[23].ToString().TrimEnd(),
                            name = row[23].ToString().TrimEnd() == "" ? null : row[23].ToString().TrimEnd(),
                            action1 = row[23].ToString().TrimEnd() == "" ? null : "noUpdate"
                        };
                        //Generates Shelf JSON from the LINQ query
                        json3 = JsonConvert.SerializeObject(query3.Skip(4), Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            DefaultValueHandling = DefaultValueHandling.Ignore
                        });
                    }

                    JArray temp = JArray.Parse(json);
                    //remove blank row
                    while (temp.Last.Count() == 0)
                    {
                        temp.Remove(temp.Last);

                    }

                    JArray temp2 = JArray.Parse(json2);
                    //remove blank row
                    while (temp2.Last.Count() == 0)
                    {
                        temp2.Remove(temp2.Last);
                    }

                    JArray temp3 = JArray.Parse(json3);
                    //remove blank row
                    while (temp3.Last.Count() == 0)
                    {
                        temp3.Remove(temp3.Last);
                    }


                    //remove duplicate site objects
                    var unique2 = temp2.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp2.Count() - 1; i >= 0; i--)
                    {
                        var token = temp2[i];
                        if (!unique2.Contains(token))
                        {
                            token.Remove();
                        }
                    };

                    //remove duplicate bandwidth objects
                    var unique3 = temp3.GroupBy(x => x["name"]).Select(x => x.First()).ToList();

                    // Iterate backwards over the JObject to remove any duplicate keys
                    for (int i = temp3.Count() - 1; i >= 0; i--)
                    {
                        var token = temp3[i];
                        if (!unique3.Contains(token))
                        {
                            token.Remove();
                        }
                    };


                    //temp.Concat(temp2);
                    var test1 = new JArray((temp2.Union(temp3)).Union(temp));
                    //var arrayOfObjects = JsonConvert.SerializeObject(
                    //     new[] { temp2, temp3, temp });

                    //JArray json3 = new JArray(temp.Union(unique));
                    json = test1.ToString();
                    //Write the file to the destination path    
                    File.WriteAllText(destinationPath, json);
                }
            }
        }
    }
}