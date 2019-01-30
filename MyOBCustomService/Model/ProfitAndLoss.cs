﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MyOBCustomService.Model
{
    [Serializable]
    public class ProfitAndLoss
    {
        //
        // Summary:
        //     Name of the referenced account. (Read only)
        public string Name { get; set; }
        //
        // Summary:
        //     The DisplayID of the referenced account. (Read only)
        public string DisplayID { get; set; }
        //
        // Summary:
        //     Start Date of the reporting period
        public DateTime StartDate { get; set; }
        //
        // Summary:
        //     End Date of the reporting period
        public DateTime EndDate { get; set; }
        //
        // Summary:
      
        //
        // Summary:
        //     Year end adjustment
        //public bool YearEndAdjust { get; set; }

        public decimal AccountTotal { get; set; }
    }
}