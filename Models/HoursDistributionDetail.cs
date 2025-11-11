using System;
using System.Collections.Generic;

namespace MODSemanal.Models;

public partial class HoursDistributionDetail
{
    public int Id { get; set; }

    public int? DistributionId { get; set; }

    public string DistributionType { get; set; }

    public int? HoursAssigned { get; set; }

    public virtual ExcessHoursDistribution Distribution { get; set; }
}