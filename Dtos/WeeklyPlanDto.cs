namespace MODSemanal.Dtos
{
    public class WeeklyPlanDto
    {
        public int WeekNumber { get; set; }
        public MaterialData CuData { get; set; } = new MaterialData();
        public MaterialData AlData { get; set; } = new MaterialData();
    }
}