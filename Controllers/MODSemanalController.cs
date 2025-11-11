using ClosedXML.Excel;
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

        [HttpPut]
        [Route("UpdateWeeklyPlan/{weekNumber}")]
        public async Task<ActionResult<WeeklyPlanResponse>> UpdateWeeklyPlan(int weekNumber, [FromBody] WeeklyPlanDto request)
        {
            try
            {
                if (weekNumber != request.WeekNumber)
                {
                    return BadRequest("El número de semana en la ruta no coincide con el del cuerpo de la solicitud");
                }

                var existingPlans = await _context.WeeklyPlan
                    .Where(wp => wp.WeekNumber == request.WeekNumber)
                    .Include(wp => wp.ExcessHoursDistribution)
                        .ThenInclude(ehd => ehd.HoursDistributionDetail)
                    .ToListAsync();

                if (!existingPlans.Any())
                    return NotFound($"No se encontraron planes para la semana {request.WeekNumber}");

                if (request.CuData.ProductionVolume <= 0 || request.CuData.Mod <= 0 || request.CuData.ProductivityTarget <= 0)
                    return BadRequest("Los valores para Cobre deben ser mayores a cero");

                if (request.AlData.ProductionVolume <= 0 || request.AlData.Mod <= 0 || request.AlData.ProductivityTarget <= 0)
                    return BadRequest("Los valores para Aluminio deben ser mayores a cero");

                var planCu = existingPlans.FirstOrDefault(p => p.MaterialType == "CU");
                if (planCu != null)
                {
                    UpdateWeeklyPlanData(planCu, request.CuData);
                }

                var planAl = existingPlans.FirstOrDefault(p => p.MaterialType == "AL");
                if (planAl != null)
                {
                    UpdateWeeklyPlanData(planAl, request.AlData);
                }

                foreach (var plan in existingPlans)
                {
                    var distribution = plan.ExcessHoursDistribution.FirstOrDefault();
                    if (distribution != null)
                    {
                        UpdateExcessDistribution(distribution, plan.ExcessPersonHours.Value, plan.Mod.Value);
                    }
                }

                await _context.SaveChangesAsync();

                return Ok(new WeeklyPlanResponse
                {
                    success = true,
                    WeekNumber = request.WeekNumber,
                    PlanCreated = 0,
                    PlanUpdated = 2,
                    DistributionsUpdated = 2
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Error interno del servidor: {ex.Message}" });
            }
        }

        private void UpdateWeeklyPlanData(WeeklyPlan plan, MaterialData data)
        {
            plan.ProductivityTarget = data.ProductivityTarget;
            plan.ProductionVolume = data.ProductionVolume;
            plan.Mod = data.Mod;
            plan.HoursNeed = (int)(data.ProductionVolume / data.ProductivityTarget);
            plan.HoursPersonAvailable = (int)(data.Mod * 46.5m);
            plan.ExcessPersonHours = plan.HoursPersonAvailable - plan.HoursNeed;
            plan.ExcessHoursPerPerson = Math.Round(plan.ExcessPersonHours.Value / (decimal)data.Mod, 2);
        }

        private void UpdateExcessDistribution(ExcessHoursDistribution distribution, int totalExcessHours, int mod)
        {
            distribution.TotalExcessHours = totalExcessHours;
            distribution.Mod = mod;

            var bankHours = (int)(mod * 6.5);
            var vacationsHours = 0;
            var trainingHours = 0;
            var generalServicesHours = 0;

            distribution.TotalAvailableHours = totalExcessHours - (bankHours + vacationsHours + trainingHours + generalServicesHours);

            foreach(var detail in distribution.HoursDistributionDetail)
            {
                detail.HoursAssigned = detail.DistributionType switch
                {
                    "Banco de Horas" => bankHours,
                    "Vacaciones" => vacationsHours,
                    "Capacitación" => trainingHours,
                    "Servicios Generales" => generalServicesHours,
                    _ => detail.HoursAssigned
                };
            }
        }

        [HttpGet]
        [Route("Getall")]
        public async Task<ActionResult<IEnumerable<WeeklyPlanSummary>>> GetAll()
        {
            try
            {
                var plans = await _context.WeeklyPlan
                        .OrderByDescending(wp => wp.Id)
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


        [HttpGet]
        [Route("GenerateWeeklyReport/{weekNumber}")]
        public async Task<IActionResult> GenerateWeeklyReport(int weekNumber)
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

                var fullResponse = new WeeklyPlanFullResponse
                {
                    WeekNumber = weekNumber,
                    Plans = plans
                };

                using var workbook = new XLWorkbook();

                var summarySheet = workbook.Worksheets.Add("Resumen General");
                GenerateSummarySheet(summarySheet, fullResponse);

                var cuSheet = workbook.Worksheets.Add("Cobre (CU)");
                var cuPlan = fullResponse.Plans.FirstOrDefault(p => p.MaterialType == "CU");
                if (cuPlan != null)
                {
                    GenerateMaterialSheet(cuSheet, cuPlan, "Cobre");
                }

                var alSheet = workbook.Worksheets.Add("Aluminio (AL)");
                var alPlan = fullResponse.Plans.FirstOrDefault(p => p.MaterialType == "AL");
                if (alPlan != null)
                {
                    GenerateMaterialSheet(alSheet, alPlan, "Aluminio");
                }

                // 4. Hoja de Distribución de Horas
                var distributionSheet = workbook.Worksheets.Add("Distribución Horas");
                GenerateDistributionSheet(distributionSheet, fullResponse);

                summarySheet.Columns().AdjustToContents();
                cuSheet?.Columns().AdjustToContents();
                alSheet?.Columns().AdjustToContents();
                distributionSheet.Columns().AdjustToContents();

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                var content = stream.ToArray();

                var fileName = $"Reporte_Semana_{weekNumber}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                return File(content,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = $"Error al generar el reporte: {ex.Message}" });
            }
        }

        private void GenerateSummarySheet(IXLWorksheet sheet, WeeklyPlanFullResponse data)
        {
            sheet.Cell("A1").Value = $"REPORTE SEMANAL - SEMANA {data.WeekNumber}";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 16;
            sheet.Range("A1:E1").Merge();

            sheet.Cell("A2").Value = $"Generado el: {DateTime.Now:dd/MM/yyyy HH:mm}";
            sheet.Range("A2:E2").Merge();

            sheet.Cell("A3").Value = "";

            var headers = new[] { "Material", "Meta Productividad", "Volumen Producción", "MOD",
                             "Horas Requeridas", "Horas Disponibles", "Exceso Horas", "Exceso por Persona" };

            for (int i = 0; i < headers.Length; i++)
            {
                sheet.Cell(4, i + 1).Value = headers[i];
                sheet.Cell(4, i + 1).Style.Font.Bold = true;
                sheet.Cell(4, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                sheet.Cell(4, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int row = 5;
            foreach (var plan in data.Plans)
            {
                sheet.Cell(row, 1).Value = plan.MaterialType == "CU" ? "Cobre" : "Aluminio";
                sheet.Cell(row, 2).Value = plan.ProductivityTarget;
                sheet.Cell(row, 3).Value = plan.ProductionVolume;
                sheet.Cell(row, 4).Value = plan.Mod;
                sheet.Cell(row, 5).Value = plan.HoursNeed;
                sheet.Cell(row, 6).Value = plan.HoursPersonAvailable;
                sheet.Cell(row, 7).Value = plan.ExcessPersonHours;
                sheet.Cell(row, 8).Value = plan.ExcessHoursPerPerson;

                sheet.Cell(row, 2).Style.NumberFormat.Format = "0.00";
                sheet.Cell(row, 3).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(row, 4).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(row, 6).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(row, 7).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(row, 8).Style.NumberFormat.Format = "0.00";

                row++;
            }

            var dataRange = sheet.Range(4, 1, row - 1, headers.Length);
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            sheet.Cell(row, 1).Value = "TOTALES";
            sheet.Cell(row, 1).Style.Font.Bold = true;
            sheet.Cell(row, 3).FormulaA1 = $"=SUM(C5:C{row - 1})";
            sheet.Cell(row, 4).FormulaA1 = $"=SUM(D5:D{row - 1})";
            sheet.Cell(row, 5).FormulaA1 = $"=SUM(E5:E{row - 1})";
            sheet.Cell(row, 6).FormulaA1 = $"=SUM(F5:F{row - 1})";
            sheet.Cell(row, 7).FormulaA1 = $"=SUM(G5:G{row - 1})";

            var totalRange = sheet.Range(row, 1, row, headers.Length);
            totalRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            totalRange.Style.Font.Bold = true;
            totalRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        }

        private void GenerateMaterialSheet(IXLWorksheet sheet, WeeklyPlanWithDistribution plan, string materialName)
        {
            sheet.Cell("A1").Value = $"{materialName.ToUpper()} - SEMANA {plan.WeekNumber}";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 14;
            sheet.Range("A1:D1").Merge();

            sheet.Cell("A3").Value = "Meta de Productividad:";
            sheet.Cell("B3").Value = plan.ProductivityTarget;
            sheet.Cell("B3").Style.NumberFormat.Format = "0.00";

            sheet.Cell("A4").Value = "Volumen de Producción:";
            sheet.Cell("B4").Value = plan.ProductionVolume;
            sheet.Cell("B4").Style.NumberFormat.Format = "#,##0";

            sheet.Cell("A5").Value = "MOD:";
            sheet.Cell("B5").Value = plan.Mod;
            sheet.Cell("B5").Style.NumberFormat.Format = "#,##0";

            sheet.Cell("A6").Value = "Horas Requeridas:";
            sheet.Cell("B6").Value = plan.HoursNeed;
            sheet.Cell("B6").Style.NumberFormat.Format = "#,##0";

            sheet.Cell("A7").Value = "Horas Persona Disponibles:";
            sheet.Cell("B7").Value = plan.HoursPersonAvailable;
            sheet.Cell("B7").Style.NumberFormat.Format = "#,##0";

            sheet.Cell("A8").Value = "Exceso de Horas Persona:";
            sheet.Cell("B8").Value = plan.ExcessPersonHours;
            sheet.Cell("B8").Style.NumberFormat.Format = "#,##0";

            sheet.Cell("A9").Value = "Exceso por Persona:";
            sheet.Cell("B9").Value = plan.ExcessHoursPerPerson;
            sheet.Cell("B9").Style.NumberFormat.Format = "0.00";

            // Cálculos de eficiencia
            sheet.Cell("A11").Value = "INDICADORES DE EFICIENCIA";
            sheet.Cell("A11").Style.Font.Bold = true;
            sheet.Cell("A11").Style.Font.FontSize = 12;

            sheet.Cell("A12").Value = "Utilización de Horas:";
            sheet.Cell("B12").FormulaA1 = "=B6/B7";
            sheet.Cell("B12").Style.NumberFormat.Format = "0.00%";

            sheet.Cell("A13").Value = "Eficiencia MOD:";
            sheet.Cell("B13").FormulaA1 = "=B6/(B5*46.5)";
            sheet.Cell("B13").Style.NumberFormat.Format = "0.00%";

            var infoRange = sheet.Range("A3:B9");
            infoRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            infoRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            var efficiencyRange = sheet.Range("A12:B13");
            efficiencyRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            efficiencyRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        private void GenerateDistributionSheet(IXLWorksheet sheet, WeeklyPlanFullResponse data)
        {
            sheet.Cell("A1").Value = $"DISTRIBUCIÓN DE HORAS EXCEDENTES - SEMANA {data.WeekNumber}";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 14;
            sheet.Range("A1:F1").Merge();

            int row = 3;

            foreach (var plan in data.Plans)
            {
                var materialName = plan.MaterialType == "CU" ? "Cobre" : "Aluminio";

                sheet.Cell(row, 1).Value = $"{materialName}";
                sheet.Cell(row, 1).Style.Font.Bold = true;
                sheet.Cell(row, 1).Style.Font.FontSize = 12;
                sheet.Range($"A{row}:F{row}").Merge();
                row++;

                var distHeaders = new[] { "Concepto", "Horas Asignadas", "% del Total", "Horas por Persona" };
                for (int i = 0; i < distHeaders.Length; i++)
                {
                    sheet.Cell(row, i + 1).Value = distHeaders[i];
                    sheet.Cell(row, i + 1).Style.Font.Bold = true;
                    sheet.Cell(row, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    sheet.Cell(row, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                row++;

                int startDataRow = row;
                foreach (var detail in plan.Distribution.DistributionDetails)
                {
                    sheet.Cell(row, 1).Value = detail.Type;
                    sheet.Cell(row, 2).Value = detail.HoursAssigned;
                    sheet.Cell(row, 3).FormulaA1 = $"=B{row}/B{startDataRow + 4}";
                    sheet.Cell(row, 4).FormulaA1 = $"=B{row}/$D$2";

                    sheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0";
                    sheet.Cell(row, 3).Style.NumberFormat.Format = "0.00%";
                    sheet.Cell(row, 4).Style.NumberFormat.Format = "0.00";

                    row++;
                }

                sheet.Cell(row, 1).Value = "TOTAL DISTRIBUIDO";
                sheet.Cell(row, 1).Style.Font.Bold = true;
                sheet.Cell(row, 2).FormulaA1 = $"=SUM(B{startDataRow}:B{row - 1})";
                sheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(row, 3).Value = "";
                sheet.Cell(row, 4).Value = "";

                sheet.Cell(startDataRow, 6).Value = "Horas Excedentes Totales:";
                sheet.Cell(startDataRow + 1, 6).Value = "MOD:";
                sheet.Cell(startDataRow + 2, 6).Value = "Horas Disponibles:";

                sheet.Cell(startDataRow, 7).Value = plan.Distribution.TotalExcessHours;
                sheet.Cell(startDataRow + 1, 7).Value = plan.Mod;
                sheet.Cell(startDataRow + 2, 7).Value = plan.Distribution.TotalAvailableHours;

                sheet.Cell(startDataRow, 7).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(startDataRow + 1, 7).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(startDataRow + 2, 7).Style.NumberFormat.Format = "#,##0";

                var distRange = sheet.Range(startDataRow - 1, 1, row, 4);
                distRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                distRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                sheet.Range(row, 1, row, 4).Style.Fill.BackgroundColor = XLColor.LightBlue;

                row += 3; 
            }
        }
    }
}