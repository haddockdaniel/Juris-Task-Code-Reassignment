using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class CompositKeyRecord
    {
        public string BillNo { get; set; }
        public string Tkpr { get; set; }
        public string Matter { get; set; }
        public string TaskCode { get; set; }
        public string ActCode { get; set; }
        public string Batch { get; set; }
        public string Record { get; set; }
        public string Year { get; set; }
        public string Period { get; set; }
        public double WorkHrsBld { get; set; }
        public double HrsBld { get; set; }
        public double StdValueBld { get; set; }
        public double ActualValueBld { get; set; }
        public double ActualAmtBld { get; set; }
        public double Rcvd { get; set; }
        public double Adj { get; set; }
        public double Pend { get; set; }
        public double PrePost { get; set; }
        public double Amount { get; set; }
        public double WorkedHrsEntered { get; set; }
        public double NonBilHrsEntered { get; set; }
        public double BilHrsEntered { get; set; }
        public double FeeEnteredStdValue { get; set; }
        public double FeeEnteredActualValue { get; set; }
    }
}
