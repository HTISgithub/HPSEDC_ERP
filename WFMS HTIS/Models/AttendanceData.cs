using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models
{ 
    public class Data
    {
        public int RecordID { get; set; }
        public string Time { get; set; }
    } 
    public class AttendanceApiResponse
    {
        public string Reference { get; set; }
        public string ResponseURL { get; set; }
        public int StatusCode { get; set; }
        public string StatusString { get; set; }
        public Data Data { get; set; }
    } 
    public class Root
    {
        public AttendanceApiResponse Response { get; set; }
    }

    /**********************************************************/

    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
    public class PanoImage
    {
        public string Name { get; set; }
        public int Size { get; set; }
        public string Data { get; set; }
        public string URL { get; set; }
    }

    public class FaceImage
    {
        public string Name { get; set; }
        public int Size { get; set; }
        public string Data { get; set; }
        public string URL { get; set; }
    }

    public class FaceArea
    {
        public int LeftTopX { get; set; }
        public int LeftTopY { get; set; }
        public int RightBottomX { get; set; }
        public int RightBottomY { get; set; }
    }

    public class FaceInfoList
    {
        public int ID { get; set; }
        public double Timestamp { get; set; }
        public int CapSrc { get; set; }
        public PanoImage PanoImage { get; set; }
        public int MaskFlag { get; set; }
        public double Temperature { get; set; }
        public FaceImage FaceImage { get; set; }
        public FaceArea FaceArea { get; set; }
    }

    public class MatchPersonInfo
    {
        public string PersonCode { get; set; }
        public string PersonName { get; set; }
        public int Gender { get; set; }
        public string CardID { get; set; }
        public string IdentityNo { get; set; }
    }

    public class LibMatInfoList
    {
        public int ID { get; set; }
        public int LibID { get; set; }
        public int LibType { get; set; }
        public int MatchStatus { get; set; }
        public long MatchPersonID { get; set; }
        public long MatchFaceID { get; set; }
        public MatchPersonInfo MatchPersonInfo { get; set; }
    }

    public class ACSPassRecordList
    {
        public int FaceInfoNum { get; set; }
        public List<FaceInfoList> FaceInfoList { get; set; }
        public int CardInfoNum { get; set; }
        public List<object> CardInfoList { get; set; }
        public int GateInfoNum { get; set; }
        public List<object> GateInfoList { get; set; }
        public int LibMatInfoNum { get; set; }
        public List<LibMatInfoList> LibMatInfoList { get; set; }
    }

    public class DataLog
    {
        public int Total { get; set; }
        public int Offset { get; set; }
        public int Num { get; set; }
        public List<ACSPassRecordList> ACSPassRecordList { get; set; }
    }

    public class Response
    {
        public string ResponseURL { get; set; }
        public int CreatedID { get; set; }
        public int ResponseCode { get; set; }
        public int SubResponseCode { get; set; }
        public string ResponseString { get; set; }
        public int StatusCode { get; set; }
        public string StatusString { get; set; }
        public DataLog Data { get; set; }
    }

    public class RootLog
    {
        public Response Response { get; set; }
    }


}