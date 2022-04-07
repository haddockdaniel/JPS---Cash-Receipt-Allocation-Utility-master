using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace JurisUtilityBase
{
    public class Check
    {
        public Check()
        {
            receiptType = "";
            depDate = "";
            checkNum = "";
            checkAmt = 0.00;
            checkDate = "";
            payor = "";
            client = "";
            matter = "";
            bankCode = "";
            glAccount = "";
            reference = "";
            invDate = "";
            invNumber = 0;
            billTotal = 0.00;
            billBalance = 0.00;
            allocationAmount = 0.00;
            rowNum = 0;
            clisysnbr = 0;
            matsysnbr = 0;
            chartsysnbr = 0;
            isError = false;
            ar = 0.00;
            ppd = 0.00;
            trust = 0.00;
            noncli = 0.00;
            Hold = false;
        }

        public string receiptType { get; set; }
        [DisplayName("Deposit Date")] public string depDate { get; set; }
        [DisplayName("Check Number")] public string checkNum { get; set; }
        [DisplayName("Check Amount")] public double checkAmt { get; set; }
        [DisplayName("Check Date")] public string checkDate { get; set; }
        [DisplayName("Payor")] public string payor { get; set; }
        [DisplayName("Client")] public string client { get; set; }

        [DisplayName("Matter")] public string matter { get; set; }

        [DisplayName("Bank")] public string bankCode { get; set; }

        [DisplayName("GL Account")] public string glAccount { get; set; }

        [DisplayName("Reference")] public string reference { get; set; }

        [DisplayName("Invoice Date")] public string invDate { get; set; }

        [DisplayName("Invoice Number")] public int invNumber { get; set; }

        [DisplayName("Bill Total")] public double billTotal { get; set; }

        [DisplayName("Bill Balance")] public double billBalance { get; set; }

        [DisplayName("Allocation Amount")] public double allocationAmount { get; set; }

        public double feeAmount { get; set; }

        public double cashExpAmount { get; set; }

        public double nonCashExpAmount { get; set; }

        public double tax1Amount { get; set; }

        public double tax2Amount { get; set; }

        public double tax3Amount { get; set; }

        public double surchargeAmount { get; set; }

        public int rowNum { get; set; }

        public int clisysnbr { get; set; }
        public int matsysnbr { get; set; }

        public int chartsysnbr { get; set; }

        public bool isError { get; set; }

        public double ppd { get; set; }
        public double ar { get; set; }
        public double trust { get; set; }
        public double noncli { get; set; }

        public bool Hold { get; set; }
    }

}
