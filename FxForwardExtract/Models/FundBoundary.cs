using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FxForwardExtract.Models
{
    public class FundBoundary
    {
        public int StartLine { get; set; }
        public int EndLine { get; set; }
        public string PFPlanCode { get; set; }
        public string PFName { get; set; }
        public DateTime DateAsOf { get; set; }
    }
}
