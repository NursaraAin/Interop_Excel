using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FxForwardExtract.Models
{
    public class FXForwardData
    {
        public string PFName { get; set; }
        public string PFCode { get; set; }
        public DateTime DateAsOf { get; set; }
        public object Type { get; set; }
        public object RefNo { get; set; }
        public object Counterparty { get; set; }
        public object FromDay { get; set; }
        public object ToDay { get; set; }
        public object SoldCur { get; set; }
        public object PurchasedCur { get; set; }
        public object SoldAmount { get; set; }
        public object ContractRate { get; set; }
        public object ContractValue { get; set; }
        public object MarketRate { get; set; }
        public object MarketValue { get; set; }
        public object UGLQC { get; set; }
        public object ClosingFXRate { get; set; }
        public object UGLPC { get; set; }
    }
}
