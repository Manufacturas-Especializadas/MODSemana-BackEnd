using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using MODSemanal.Dtos;
using MODSemanal.Models;

namespace MODSemanal.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MODSemanalController : ControllerBase
    {
        private readonly AppDbContext _context;

        public MODSemanalController(AppDbContext context)
        {
            _context = context;
        }

        [HttpPost]
        [Route("CreateWeeklyPlan")]
        public async Task<ActionResult<WeeklyPlanResponse>> CreateWeeklyPlan(WeeklyPlanDto request)
        {
            try 
            { 
                if (request.CuData.ProductionVolume <= 0 || request.CuData.Mod <= 0 || request.CuData.ProductivityTarget <= 0)
                    return BadRequest("Los valores para Cobre deben ser mayores a cero");

                if (request.AlData.ProductionVolume <= 0 || request.AlData.Mod <= 0 || request.AlData.ProductivityTarget <= 0)
                    return BadRequest("Los valores para Aluminio deben ser mayores a cero");

                var weeklyPlans = new List<WeeklyPlan>();

                var planCu = CalculateWeeklyPlan(request.WeekNumber, "CU", request.CuData);
                weeklyPlans.Add(planCu);

                var planAl = CalculateWeeklyPlan(request.WeekNumber, "AL", request.AlData);
                weeklyPlans.Add(planAl);

                _context.WeeklyPlan.AddRange(weeklyPlans);
                await _context.SaveChangesAsync();

                var distributions = new List<ExcessHoursDistribution>();

                var distributionCu = CalculateExcessDistribution(planCu.Id, "CU", planCu.ExcessPersonHours.Value, planCu.Mod.Value);
                distributions.Add(distributionCu);

                var distributionAl = CalculateExcessDistribution(planAl.Id, "AL", planAl.ExcessPersonHours.Value, planAl.Mod.Value);
                distributions.Add(distributionAl);

                _context.ExcessHoursDistribution.AddRange(distributions);
                await _context.SaveChangesAsync();

                return Ok(new WeeklyPlanResponse
                {
                    success = true,
                    WeekNumber = request.WeekNumber,
                    PlanCreated= 2,
                    DistributionsCreated = 2
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Error interno del servidor: {ex.Message}" });
            }
        }

        private WeeklyPlan CalculateWeeklyPlan(int weekNumber, string materialType, MaterialData data)
        {
            var hoursNeed = (int)(data.ProductionVolume / data.ProductivityTarget);
            var hoursPersonAvailable = (int)(data.Mod * 46.5m);
            var excessPersonHours = hoursPersonAvailable - hoursNeed;
            var excessHoursPerPerson = excessPersonHours / (decimal)data.Mod;

            return new WeeklyPlan
            {
                WeekNumber = weekNumber,
                MaterialType = materialType,
                ProductivityTarget = data.ProductivityTarget,
                ProductionVolume = data.ProductionVolume,
                Mod = data.Mod,
                HoursNeed = hoursNeed,
                HoursPersonAvailable = hoursPersonAvailable,
                ExcessPersonHours = excessPersonHours,
                ExcessHoursPerPerson = Math.Round(excessHoursPerPerson, 2)
            };
        }

        private ExcessHoursDistribution CalculateExcessDistribution(int weeklyPlanId, string materialType, int totalExcessHours, int mod)
        {
            var bankHours = (int)(mod * 6.5m);
            var vacationsHours = 0;
            var trainingHours = 0;
            var generalServicesHours = 0;

            var totalAvailableHours = totalExcessHours - (bankHours + vacationsHours + trainingHours + trainingHours);

            var distribution = new ExcessHoursDistribution
            {
                WeeklyPlanId = weeklyPlanId,
                MaterialType = materialType,
                TotalExcessHours = totalExcessHours,
                Mod = mod,
                TotalAvailableHours = totalAvailableHours,
            };

            distribution.HoursDistributionDetail = new List<HoursDistributionDetail>
            {
                new HoursDistributionDetail { DistributionType = "Vacaciones", HoursAssigned = vacationsHours },
                new HoursDistributionDetail { DistributionType = "Banco de Horas", HoursAssigned = bankHours },
                new HoursDistributionDetail { DistributionType = "Capacitación", HoursAssigned = trainingHours },
                new HoursDistributionDetail { DistributionType = "Servicios Generales", HoursAssigned = generalServicesHours }
            };
            
            return distribution;
        }

        [HttpGet]
        [Route("Getall")]
        public async Task<ActionResult<IEnumerable<WeeklyPlanSummary>>> GetAll()
        {
            try
            {
                var plans = await _context.WeeklyPlan
                        .Select(wp => new WeeklyPlanSummary
                        {
                            Id = wp.Id,
                            WeekNumber = wp.WeekNumber,
                            MaterialType = wp.MaterialType,
                            ProductivityTarget = wp.ProductivityTarget,
                            ProductionVolume = wp.ProductionVolume,
                            HoursNeed = wp.HoursNeed,
                            Mod = wp.Mod,
                            HoursPersonAvailable = wp.HoursPersonAvailable,
                            ExcessPersonHours = wp.ExcessPersonHours,
                            ExcessHoursPerPerson = wp.ExcessHoursPerPerson
                        })
                        .AsNoTracking()
                        .ToListAsync();

                return Ok(plans);
            }
            catch(Exception ex)
            {
                return StatusCode(500, new { error = $"Error al obtener datos: {ex.Message}" });
            }
        }

        [HttpGet]
        [Route("GetByWeek/{weekNumber}")]
        public async Task<ActionResult<WeeklyPlanFullResponse>> GetFullByWeek(int weekNumber)
        {
            try
            {
                var plans = await _context.WeeklyPlan
                    .Where(wp => wp.WeekNumber == weekNumber)
                    .Include(wp => wp.ExcessHoursDistribution)
                        .ThenInclude(ehd => ehd.HoursDistributionDetail)
                    .Select(wp => new WeeklyPlanWithDistribution
                    {
                        Id = wp.Id,
                        WeekNumber = wp.WeekNumber,
                        MaterialType = wp.MaterialType,
                        ProductivityTarget = wp.ProductivityTarget,
                        ProductionVolume = wp.ProductionVolume,
                        HoursNeed = wp.HoursNeed,
                        Mod = wp.Mod,
                        HoursPersonAvailable = wp.HoursPersonAvailable,
                        ExcessPersonHours = wp.ExcessPersonHours,
                        ExcessHoursPerPerson = wp.ExcessHoursPerPerson,
                        Distribution = wp.ExcessHoursDistribution.Select(ehd => new DistributionSummary
                        {
                            Id = ehd.Id,
                            TotalExcessHours = ehd.TotalExcessHours,
                            Mod = ehd.Mod,
                            TotalAvailableHours = ehd.TotalAvailableHours,
                            DistributionDetails = ehd.HoursDistributionDetail.Select(hdd => new DistributionDetailSummary
                            {
                                Type = hdd.DistributionType,
                                HoursAssigned = hdd.HoursAssigned
                            }).ToList()
                        }).FirstOrDefault()!
                    })
                    .AsNoTracking()
                    .ToListAsync();

                if (!plans.Any())
                    return NotFound($"No se encontraron planes para la semana {weekNumber}");

                return Ok(new WeeklyPlanFullResponse
                {
                    WeekNumber = weekNumber,
                    Plans = plans
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Error al obtener datos: {ex.Message}" });
            }
        }
    }
}