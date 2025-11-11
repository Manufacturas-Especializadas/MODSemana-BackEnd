using System;
using System.Collections.Generic;

namespace MODSemanal.Models;

public partial class ExcessHoursDistribution
{
    public int Id { get; set; }

    public int? WeeklyPlanId { get; set; }

    public string MaterialType { get; set; }

    public int? TotalExcessHours { get; set; }

    public int? Mod { get; set; }

    public int? TotalAvailableHours { get; set; }

    public virtual ICollection<HoursDistributionDetail> HoursDistributionDetail { get; set; } = new List<HoursDistributionDetail>();

    public virtual WeeklyPlan WeeklyPlan { get; set; }
}