namespace MODSemanal.Dtos
{
    public class WeeklyPlanFullResponse
    {
        public int WeekNumber { get; set; }

        public List<WeeklyPlanWithDistribution> Plans { get; set; } = new List<WeeklyPlanWithDistribution>();
    }
}