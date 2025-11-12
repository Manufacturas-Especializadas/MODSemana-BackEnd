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
                GenerateProfessionalSummarySheet(summarySheet, fullResponse);

                var cuSheet = workbook.Worksheets.Add("Cobre (CU)");
                var cuPlan = fullResponse.Plans.FirstOrDefault(p => p.MaterialType == "CU");
                if (cuPlan != null)
                {
                    GenerateMaterialSheetWithoutEfficiency(cuSheet, cuPlan, "Cobre");
                }

                var alSheet = workbook.Worksheets.Add("Aluminio (AL)");
                var alPlan = fullResponse.Plans.FirstOrDefault(p => p.MaterialType == "AL");
                if (alPlan != null)
                {
                    GenerateMaterialSheetWithoutEfficiency(alSheet, alPlan, "Aluminio");
                }

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

        private void GenerateProfessionalSummarySheet(IXLWorksheet sheet, WeeklyPlanFullResponse data)
        {
            sheet.Cell("A1").Value = $"PLAN MOD PARA SEMANA COBRE Y ALUMINIO";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 16;
            sheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range("A1:I1").Merge();

            sheet.Cell("A2").Value = $"SEMANA # {data.WeekNumber}";
            sheet.Cell("A2").Style.Font.Bold = true;
            sheet.Cell("A2").Style.Font.FontSize = 12;
            sheet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range("A2:I2").Merge();

            sheet.Cell("A3").Value = $"Generado: {DateTime.Now:dd/MM/yyyy HH:mm}";
            sheet.Cell("A3").Style.Font.Italic = true;
            sheet.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range("A3:I3").Merge();

            var mainHeaders = new[] { "", "CU", "AL", "CONCEPTO", "TOTAL" };
            int headerRow = 5;

            for (int i = 0; i < mainHeaders.Length; i++)
            {
                var cell = sheet.Cell(headerRow, i + 1);
                cell.Value = mainHeaders[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int dataRow = 6;
            var cuPlan = data.Plans.First(p => p.MaterialType == "CU");
            var alPlan = data.Plans.First(p => p.MaterialType == "AL");

            sheet.Cell(dataRow, 1).Value = "Objetivo Productividad Cu y AL.";
            sheet.Cell(dataRow, 2).Value = cuPlan.ProductivityTarget;
            sheet.Cell(dataRow, 3).Value = alPlan.ProductivityTarget;
            sheet.Cell(dataRow, 4).Value = "Kgs/1tr";
            sheet.Cell(dataRow, 5).Value = "";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "Volumen a Fabricar/Semana";
            sheet.Cell(dataRow, 2).Value = cuPlan.ProductionVolume;
            sheet.Cell(dataRow, 3).Value = alPlan.ProductionVolume;
            sheet.Cell(dataRow, 4).Value = "Kgs";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "Horas Necesarias";
            sheet.Cell(dataRow, 2).Value = cuPlan.HoursNeed;
            sheet.Cell(dataRow, 3).Value = alPlan.HoursNeed;
            sheet.Cell(dataRow, 4).Value = "Horas Persona";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "MOD";
            sheet.Cell(dataRow, 2).Value = cuPlan.Mod;
            sheet.Cell(dataRow, 3).Value = alPlan.Mod;
            sheet.Cell(dataRow, 4).Value = "Personas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "Horas Persona Disponible";
            sheet.Cell(dataRow, 2).Value = cuPlan.HoursPersonAvailable;
            sheet.Cell(dataRow, 3).Value = alPlan.HoursPersonAvailable;
            sheet.Cell(dataRow, 4).Value = "Horas Persona";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "Excedente Horas Persona";
            sheet.Cell(dataRow, 2).Value = cuPlan.ExcessPersonHours;
            sheet.Cell(dataRow, 3).Value = alPlan.ExcessPersonHours;
            sheet.Cell(dataRow, 4).Value = "Horas Totales";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "Excedente Horas / Persona";
            sheet.Cell(dataRow, 2).Value = cuPlan.ExcessHoursPerPerson;
            sheet.Cell(dataRow, 3).Value = alPlan.ExcessHoursPerPerson;
            sheet.Cell(dataRow, 4).Value = "Horas / Persona";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=E{dataRow - 1}/D{dataRow - 3}";
            dataRow++;

            for (int row = 6; row <= dataRow; row++)
            {
                sheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00";
                sheet.Cell(row, 3).Style.NumberFormat.Format = "#,##0.00";

                if (row > 6 && row < dataRow - 1)
                {
                    sheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
                }
                else if (row == dataRow - 1) 
                {
                    sheet.Cell(row, 5).Style.NumberFormat.Format = "0.00";
                }
            }

            var tableRange = sheet.Range(headerRow, 1, dataRow - 1, 5);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            var labelRange = sheet.Range(6, 1, dataRow - 1, 1);
            labelRange.Style.Font.Bold = true;
            labelRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            var valuesRange = sheet.Range(6, 2, dataRow - 1, 3);
            valuesRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            var conceptsRange = sheet.Range(6, 4, dataRow - 1, 4);
            conceptsRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            var totalsRange = sheet.Range(6, 5, dataRow - 1, 5);
            totalsRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            totalsRange.Style.Font.Bold = true;
        }

        private void GenerateMaterialSheetWithoutEfficiency(IXLWorksheet sheet, WeeklyPlanWithDistribution plan, string materialName)
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

            var infoRange = sheet.Range("A3:B9");
            infoRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            infoRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            infoRange.Style.Fill.BackgroundColor = XLColor.White;
            for (int i = 3; i <= 9; i++)
            {
                sheet.Cell(i, 1).Style.Font.Bold = true;
                sheet.Cell(i, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
            }
        }

        private void GenerateDistributionSheet(IXLWorksheet sheet, WeeklyPlanFullResponse data)
        {
            sheet.Cell("A1").Value = $"DISTRIBUCIÓN DE EXCEDENTE HORAS/COLABORADOR";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").Style.Font.FontSize = 14;
            sheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range("A1:E1").Merge();

            sheet.Cell("A2").Value = $"SEMANA {data.WeekNumber}";
            sheet.Cell("A2").Style.Font.Bold = true;
            sheet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range("A2:E2").Merge();

            var mainHeaders = new[] { "", "CU", "AL", "CONCEPTO" };
            int headerRow = 4;

            for (int i = 0; i < mainHeaders.Length; i++)
            {
                var cell = sheet.Cell(headerRow, i + 1);
                cell.Value = mainHeaders[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            sheet.Range(headerRow, 5, headerRow, 5).Merge();
            sheet.Cell(headerRow, 5).Value = "TOTAL";
            sheet.Cell(headerRow, 5).Style.Font.Bold = true;
            sheet.Cell(headerRow, 5).Style.Fill.BackgroundColor = XLColor.LightGray;
            sheet.Cell(headerRow, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell(headerRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            int dataRow = 5;
            var cuPlan = data.Plans.First(p => p.MaterialType == "CU");
            var alPlan = data.Plans.First(p => p.MaterialType == "AL");
            var cuDistribution = cuPlan.Distribution;
            var alDistribution = alPlan.Distribution;

            sheet.Cell(dataRow, 1).Value = "Total Excedente Horas";
            sheet.Cell(dataRow, 2).Value = cuDistribution.TotalExcessHours;
            sheet.Cell(dataRow, 3).Value = alDistribution.TotalExcessHours;
            sheet.Cell(dataRow, 4).Value = "Horas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "MOD";
            sheet.Cell(dataRow, 2).Value = cuDistribution.Mod;
            sheet.Cell(dataRow, 3).Value = alDistribution.Mod;
            sheet.Cell(dataRow, 4).Value = "Personas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            var vacacionesCu = cuDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Vacaciones")?.HoursAssigned ?? 0;
            var vacacionesAl = alDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Vacaciones")?.HoursAssigned ?? 0;

            var bancoHorasCu = cuDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Banco de Horas")?.HoursAssigned ?? 0;
            var bancoHorasAl = alDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Banco de Horas")?.HoursAssigned ?? 0;

            var capacitacionCu = cuDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Capacitación")?.HoursAssigned ?? 0;
            var capacitacionAl = alDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Capacitación")?.HoursAssigned ?? 0;

            var serviciosCu = cuDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Servicios Generales")?.HoursAssigned ?? 0;
            var serviciosAl = alDistribution.DistributionDetails.FirstOrDefault(d => d.Type == "Servicios Generales")?.HoursAssigned ?? 0;

            sheet.Cell(dataRow, 1).Value = "-Vacaciones";
            sheet.Cell(dataRow, 2).Value = vacacionesCu;
            sheet.Cell(dataRow, 3).Value = vacacionesAl;
            sheet.Cell(dataRow, 4).Value = "Horas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "-Banco de Horas (6.5Hrs/Persona)";
            sheet.Cell(dataRow, 2).Value = bancoHorasCu;
            sheet.Cell(dataRow, 3).Value = bancoHorasAl;
            sheet.Cell(dataRow, 4).Value = "Horas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "-Capacitación";
            sheet.Cell(dataRow, 2).Value = capacitacionCu;
            sheet.Cell(dataRow, 3).Value = capacitacionAl;
            sheet.Cell(dataRow, 4).Value = "Horas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "-Servicios Generales";
            sheet.Cell(dataRow, 2).Value = serviciosCu;
            sheet.Cell(dataRow, 3).Value = serviciosAl;
            sheet.Cell(dataRow, 4).Value = "Horas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            sheet.Cell(dataRow, 1).Value = "=Total Horas Disponibles";
            sheet.Cell(dataRow, 2).Value = cuDistribution.TotalAvailableHours;
            sheet.Cell(dataRow, 3).Value = alDistribution.TotalAvailableHours;
            sheet.Cell(dataRow, 4).Value = "Horas";
            sheet.Cell(dataRow, 5).FormulaA1 = $"=B{dataRow}+C{dataRow}";
            dataRow++;

            for (int row = 5; row <= dataRow; row++)
            {
                sheet.Cell(row, 2).Style.NumberFormat.Format = "#,##0";
                sheet.Cell(row, 3).Style.NumberFormat.Format = "#,##0";

                sheet.Cell(row, 5).Style.NumberFormat.Format = "#,##0";
            }

            var tableRange = sheet.Range(headerRow, 1, dataRow - 1, 5);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            var labelRange = sheet.Range(5, 1, dataRow - 1, 1);
            labelRange.Style.Font.Bold = true;
            labelRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            var valuesRange = sheet.Range(5, 2, dataRow - 1, 3);
            valuesRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            var conceptsRange = sheet.Range(5, 4, dataRow - 1, 4);
            conceptsRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;


            var totalsRange = sheet.Range(5, 5, dataRow - 1, 5);
            totalsRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            totalsRange.Style.Font.Bold = true;

            sheet.Range(5, 1, 5, 5).Style.Fill.BackgroundColor = XLColor.LightBlue; 
            sheet.Range(6, 1, 6, 5).Style.Fill.BackgroundColor = XLColor.LightBlue;
            sheet.Range(dataRow - 1, 1, dataRow - 1, 5).Style.Fill.BackgroundColor = XLColor.LightGreen;

            sheet.Cell(dataRow + 1, 2).Value = cuDistribution.TotalExcessHours + alDistribution.TotalExcessHours;
            sheet.Cell(dataRow + 1, 2).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(dataRow + 1, 2).Style.Font.Bold = true;

            sheet.Cell(dataRow + 2, 2).Value = cuDistribution.Mod + alDistribution.Mod;
            sheet.Cell(dataRow + 2, 2).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(dataRow + 2, 2).Style.Font.Bold = true;

            sheet.Cell(dataRow + 3, 2).Value = (bancoHorasCu + bancoHorasAl) + (vacacionesCu + vacacionesAl) + (capacitacionCu + capacitacionAl) + (serviciosCu + serviciosAl);
            sheet.Cell(dataRow + 3, 2).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(dataRow + 3, 2).Style.Font.Bold = true;

            sheet.Cell(dataRow + 4, 2).Value = cuDistribution.TotalAvailableHours + alDistribution.TotalAvailableHours;
            sheet.Cell(dataRow + 4, 2).Style.NumberFormat.Format = "#,##0";
            sheet.Cell(dataRow + 4, 2).Style.Font.Bold = true;
        }
    }
}