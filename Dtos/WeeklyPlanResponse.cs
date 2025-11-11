namespace MODSemanal.Dtos
{
    public class WeeklyPlanResponse
    {
        public bool success {  get; set; }

        public int WeekNumber { get; set; }

        public int PlanCreated { get; set; }

        public int DistributionsCreated { get; set; }
    }
}