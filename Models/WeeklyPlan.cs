using System;
using System.Collections.Generic;

namespace MODSemanal.Models;

public partial class WeeklyPlan
{
    public int Id { get; set; }

    public int? WeekNumber { get; set; }

    public string MaterialType { get; set; }

    public decimal? ProductivityTarget { get; set; }

    public int? ProductionVolume { get; set; }

    public int? HoursNeed { get; set; }

    public int? Mod { get; set; }

    public int? HoursPersonAvailable { get; set; }

    public int? ExcessPersonHours { get; set; }

    public decimal? ExcessHoursPerPerson { get; set; }

    public virtual ICollection<ExcessHoursDistribution> ExcessHoursDistribution { get; set; } = new List<ExcessHoursDistribution>();
}