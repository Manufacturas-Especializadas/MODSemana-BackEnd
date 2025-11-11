namespace MODSemanal.Dtos
{
    public class DistributionSummary
    {
        public int Id { get; set; }
        public int? TotalExcessHours { get; set; }
        public int? Mod { get; set; }
        public int? TotalAvailableHours { get; set; }
        public List<DistributionDetailSummary> DistributionDetails { get; set; } = new List<DistributionDetailSummary>();
    }
}