using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Payroll.portal.Models
{
    //public class InvoiceData
    //{
    //    public string Version { get; set; }
    //    public List<TranDtls> Trandtls { get; set; }
    //     public List<DocDtls> DocDtls { get; set; }
    //}
    //public class TranDtls
    //{
    //    public string Taxsch { get; set; }
    //    public string Suptyp { get; set; }
    //    public string Regrev { get; set; }
    //    public string Ecmgstin { get; set; }
    //    public string Igstonintra { get; set; }
    //}
    //public class DocDtls
    //{
    //    public string Typ { get; set; }
    //    public string No { get; set; }
    //    public string Dt { get; set; }

    //}

    //public class TranDtls
    //{
    //    public string TaxSch { get; set; }
    //    public string SupTyp { get; set; }
    //    public string RegRev { get; set; }
    //    public object EcmGstin { get; set; }
    //    public string IgstOnIntra { get; set; }
    //}

    //public class DocDtls
    //{
    //    public string Typ { get; set; }
    //    public string No { get; set; }
    //    public string Dt { get; set; }
    //}

    //public class BuyerDtls
    //{
    //    public string Gstin { get; set; }
    //    public string LglNm { get; set; }
    //    public string TrdNm { get; set; }
    //    public string Pos { get; set; }
    //    public string Addr1 { get; set; }
    //    public string Addr2 { get; set; }
    //    public string Loc { get; set; }
    //    public string Pin { get; set; }
    //    public string Stcd { get; set; }
    //    public string Ph { get; set; }
    //    public string Em { get; set; }
    //}

    //public class SellerDtls
    //{
    //    public string Gstin { get; set; }
    //    public string LglNm { get; set; }
    //    public string TrdNm { get; set; }
    //    public string Addr1 { get; set; }
    //    public string Addr2 { get; set; }
    //    public string Loc { get; set; }
    //    public string Pin { get; set; }
    //    public string Stcd { get; set; }
    //    public string Ph { get; set; }
    //    public string Em { get; set; }
    //}

    //public class DispDtls
    //{
    //    public string Nm { get; set; }
    //    public string Addr1 { get; set; }
    //    public string Addr2 { get; set; }
    //    public string Loc { get; set; }
    //    public string Pin { get; set; }
    //    public string Stcd { get; set; }
    //}

    //public class ShipDtls
    //{
    //    public string Gstin { get; set; }
    //    public string LglNm { get; set; }
    //    public string TrdNm { get; set; }
    //    public string Addr1 { get; set; }
    //    public string Addr2 { get; set; }
    //    public string Loc { get; set; }
    //    public string Pin { get; set; }
    //    public string Stcd { get; set; }
    //}

    //public class EwbDtls
    //{
    //    public string TransId { get; set; }
    //    public string TransName { get; set; }
    //    public string TransMode { get; set; }
    //    public string Distance { get; set; }
    //    public string TransDocNo { get; set; }
    //    public string TransDocDt { get; set; }
    //    public string VehNo { get; set; }
    //    public string VehType { get; set; }
    //}

    //public class ExpDtls
    //{
    //    public string ShipBNo { get; set; }
    //    public string ShipBDt { get; set; }
    //    public string CntCode { get; set; }
    //    public string ForCur { get; set; }
    //    public string Port { get; set; }
    //    public string RefClm { get; set; }
    //    public string ExpDuty { get; set; }
    //}

    //public class AttribDtl
    //{
    //    public string Nm { get; set; }
    //    public string Val { get; set; }
    //}

    //public class BchDtls
    //{
    //    public string Nm { get; set; }
    //    public string ExpDt { get; set; }
    //    public string WrDt { get; set; }
    //}

    //public class ItemList
    //{
    //    // public List<AttribDtl> AttribDtls { get; set; }
    //    public string PrdSlNo { get; set; }
    //    public string OrgCntry { get; set; }
    //    public string OrdLineRef { get; set; }
    //    public double TotItemVal { get; set; }
    //    public double OthChrg { get; set; }
    //    public double StateCesNonAdvlAmt { get; set; }
    //    public double StateCesAmt { get; set; }
    //    public double StateCesRt { get; set; }
    //    public double CesNonAdvlAmt { get; set; }
    //    public double CesAmt { get; set; }
    //    public double CesRt { get; set; }
    //    public double SgstAmt { get; set; }
    //    public double CgstAmt { get; set; }
    //    public double IgstAmt { get; set; }
    //    public double Qty { get; set; }
    //    public double AssAmt { get; set; }
    //    public double PreTaxVal { get; set; }
    //    public double Discount { get; set; }
    //    public double TotAmt { get; set; }
    //    public double UnitPrice { get; set; }
    //    public string Unit { get; set; }
    //    public double FreeQty { get; set; }
    //    public double GstRt { get; set; }
    //    public string Barcde { get; set; }
    //    public string HsnCd { get; set; }
    //    public string IsServc { get; set; }
    //    public string PrdDesc { get; set; }
    //    public string SlNo { get; set; }
    //}

    //public class ValDtls
    //{
    //    public double AssVal { get; set; }
    //    public double CgstVal { get; set; }
    //    public double SgstVal { get; set; }
    //    public double IgstVal { get; set; }
    //    public double CesVal { get; set; }
    //    public double StCesVal { get; set; }
    //    public double RndOffAmt { get; set; }
    //    public double TotInvVal { get; set; }
    //    public double TotInvValFc { get; set; }
    //    public double Discount { get; set; }
    //    public double OthChrg { get; set; }
    //}

    //public class PayDtls
    //{
    //    public string Nm { get; set; }
    //    public string AccDet { get; set; }
    //    public string Mode { get; set; }
    //    public string FinInsBr { get; set; }
    //    public string CrTrn { get; set; }
    //    public string PayInstr { get; set; }
    //    public string PayTerm { get; set; }
    //    public string DirDr { get; set; }
    //    public string CrDay { get; set; }
    //    public string PaidAmt { get; set; }
    //    public string PaymtDue { get; set; }
    //}

    //public class PrecDocDtl
    //{
    //    public string InvNo { get; set; }
    //    public string InvDt { get; set; }
    //    public string OthRefNo { get; set; }
    //}

    //public class ContrDtls
    //{
    //    public string RecAdvDt { get; set; }
    //    public string RecAdvRefr { get; set; }
    //    public string TendRefr { get; set; }
    //    public string ContrRefr { get; set; }
    //    public string ExtRefr { get; set; }
    //    public string ProjRefr { get; set; }
    //    public string PORefr { get; set; }
    //    public string PORefDt { get; set; }
    //}

    //public class DocPerdDtls
    //{
    //    public string InvStDt { get; set; }
    //    public string InvEndDt { get; set; }
    //}

    //public class RefDtls
    //{
    //    public string InvRm { get; set; }
    //    //public List<PrecDocDtl> PrecDocDtls { get; set; }
    //    public List<ContrDtls> ContrDtls { get; set; }
    //    public DocPerdDtls DocPerdDtls { get; set; }
    //}

    //public class AddlDocDtl
    //{
    //    public string Url { get; set; }
    //    public string Docs { get; set; }
    //    public string Info { get; set; }
    //}

    //public class Invoicedata
    //{
    //    public string Version { get; set; }
    //    public TranDtls TranDtls { get; set; }
    //    public DocDtls DocDtls { get; set; }
    //    public BuyerDtls BuyerDtls { get; set; }
    //    public SellerDtls SellerDtls { get; set; }
    //    public DispDtls DispDtls { get; set; }
    //    public ShipDtls ShipDtls { get; set; }
    //    public BchDtls BchDtls { get; set; }
    //    public EwbDtls EwbDtls { get; set; }
    //    public ExpDtls ExpDtls { get; set; }
    //    public List<ItemList> ItemList { get; set; }
    //    public ValDtls ValDtls { get; set; }
    //    public PayDtls PayDtls { get; set; }
    //    public RefDtls RefDtls { get; set; }
    // public List<ContrDtls> ContrDtls { get; set; }
    //    public PrecDocDtl PrecDocDtl { get; set; }
    //    public AddlDocDtl AddlDocDtls { get; set; }
    //    public List<AttribDtl> AttribDtl { get;set;}
    //}

    public class PayDtls
    {
        public string Nm { get; set; }
        public string PayInstr { get; set; }
        public string FinInsBr { get; set; }
        public string Mode { get; set; }
        public string PayTerm { get; set; }
        public string AccDet { get; set; }
    }

    public class ShipDtls
    {
        public string TrdNm { get; set; }
        public string Loc { get; set; }
        public int Pin { get; set; }
        public string Addr1 { get; set; }
        public string Gstin { get; set; }
        public string Addr2 { get; set; }
        public string LglNm { get; set; }
        public string Stcd { get; set; }
    }

    public class DocDtls
    {
        public string Typ { get; set; }
        public string No { get; set; }
        public string Dt { get; set; }
    }

    public class ContrDtl
    {
        public string porefr { get; set; }
        public string porefDt { get; set; }
        public string PORefr { get; set; }
        public string PORefDt { get; set; }
    }

    public class RefDtls
    {
        public List<ContrDtl> ContrDtls { get; set; }
    }

    public class SellerDtls
    {
        public string Em { get; set; }
        public string TrdNm { get; set; }
        public string Loc { get; set; }
        public int Pin { get; set; }
        public string Addr1 { get; set; }
        public string Gstin { get; set; }
        public string Addr2 { get; set; }
        public string Ph { get; set; }
        public string LglNm { get; set; }
        public string Stcd { get; set; }
    }

    public class BuyerDtls
    {
        public string TrdNm { get; set; }
        public string Loc { get; set; }
        public int Pin { get; set; }
        public string Addr1 { get; set; }
        public string Gstin { get; set; }
        public string Addr2 { get; set; }
        public string Pos { get; set; }
        public string LglNm { get; set; }
        public string Stcd { get; set; }
    }

    public class ItemList
    {
        public double CesNonAdvlAmt { get; set; }
        public double TotItemVal { get; set; }
        public double SgstAmt { get; set; }
        public double UnitPrice { get; set; }
        public double CgstAmt { get; set; }
        public string HsnCd { get; set; }
        public string IsServc { get; set; }
        public double IgstAmt { get; set; }
        public string SlNo { get; set; }
        public double StateCesAmt { get; set; }
        public double CgstRt { get; set; }
        public double CesAmt { get; set; }
        public string PrdDesc { get; set; }
        public double StateCesRt { get; set; }
        public double AssAmt { get; set; }
        public double SgstRt { get; set; }
        public double IgstRt { get; set; }
        public double GstRt { get; set; }
        public string Unit { get; set; }
        public double TotAmt { get; set; }
        public double CesRt { get; set; }
        public double Qty { get; set; }
    }

    public class TranDtls
    {
        public string TaxSch { get; set; }
        public string SupTyp { get; set; }
        public string RegRev { get; set; }
        public string IgstOnIntra { get; set; }
    }

    public class ValDtls
    {
        public double AssVal { get; set; }
        public double StCesVal { get; set; }
        public double IgstVal { get; set; }
        public double CesVal { get; set; }
        public double TotInvVal { get; set; }
        public double CgstVal { get; set; }
        public double SgstVal { get; set; }
        public double TotInvValFc { get; set; }
    }

    public class DispDtls
    {
        public string Nm { get; set; }
        public string Loc { get; set; }
        public int Pin { get; set; }
        public string Addr1 { get; set; }
        public string Addr2 { get; set; }
        public string Stcd { get; set; }
    }

    public class ExpDtls
    {
        public string RefClm { get; set; }
        public string ShipBNo { get; set; }
        public string ShipBDt { get; set; }
        public string CntCode { get; set; }
        public string Port { get; set; }
        public string ForCur { get; set; }
    }

    public class Invoicedata
    {
        public string Version { get; set; }
        public PayDtls PayDtls { get; set; }
        public ShipDtls ShipDtls { get; set; }
        public string id { get; set; }
        public DocDtls DocDtls { get; set; }
        public RefDtls RefDtls { get; set; }
        public SellerDtls SellerDtls { get; set; }
        public BuyerDtls BuyerDtls { get; set; }
        public List<ItemList> ItemList { get; set; }
        public TranDtls TranDtls { get; set; }
        public ValDtls ValDtls { get; set; }
        public DispDtls DispDtls { get; set; }
        public ExpDtls ExpDtls { get; set; }
    }
    public class EWAYBill
    {
        public string Irn { get; set; }
        public string TransId { get; set; }
        public string TransMode { get; set; }
        public string TrnDocNO { get; set; }
        public string TrnDocDt { get; set; }
        public string VehNo { get; set; }
        public int    Distance { get; set; }
        public string VehType { get; set; }
        public string TransName { get; set; }
    }

}