using OfficeOpenXml;
using OfficeOpenXml.Style; // Para estilos básicos si los necesitamos
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ReadAndConsolidateExcel
{
    public class ExcelDataWriter
    {
        // Definimos las cabeceras aquí para fácil referencia y consistencia
        private static readonly string[] Headers = new string[] {
            "PERIODO", "RUT", "APELLIDO PATERNO", "APELLIDO MATERNO", "NOMBRES",
            "SUELDO BASE", "CENTRO DE COSTO", "DÍAS TRABAJADOS", "ATRASO", "VACACIONES",
            "ISAPRE/FONASA", "PLAN", "AFP", "%AFP", "SUELDO MENSUAL", "GRATIFICACIÓN",
            "TOTAL IMPONIBLE", "C. FAMILIAR", "TOTAL NO IMPONIBLE", "AFP", // Monto AFP
            "APV1", "APV2", "SALUD", // Monto Salud
            "S. CESANTIA", // Monto S. Cesantía
            "I.U.", "TOTAL DESCUENTOS", "TOTAL O. DESCUENTOS", "LIQUIDO A PAGAR",
            "SIS", "MUTUAL", "S. CESANTIA", // Aporte Empleador S. Cesantía (si es diferente)
            "TRIBUTABLE"
        };

        public bool EscribirConsolidado(List<LiquidacionData> datosParaEscribir, string rutaArchivoDestino, string anioSeleccionado)
        {
            if (datosParaEscribir == null || !datosParaEscribir.Any())
            {
                Console.WriteLine("No hay datos para escribir en el archivo de destino.");
                return false;
            }

            FileInfo fileInfo = new FileInfo(rutaArchivoDestino);

            // Configurar el contexto de licencia para EPPlus si es necesario
            // ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // o LicenseContext.Commercial

            try
            {
                using (var package = new ExcelPackage(fileInfo))
                {
                    // Obtener la hoja correspondiente al año, o crearla si no existe
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == anioSeleccionado);
                    if (worksheet == null)
                    {
                        worksheet = package.Workbook.Worksheets.Add(anioSeleccionado);
                    }

                    int FilaParaEscribir = 1; // Por defecto, empezamos en la fila 1

                    // Escribir encabezados si la hoja está vacía (o es nueva)
                    if (worksheet.Dimension == null || worksheet.Dimension.End.Row == 0)
                    {
                        for (int i = 0; i < Headers.Length; i++)
                        {
                            worksheet.Cells[FilaParaEscribir, i + 1].Value = Headers[i];
                            // Opcional: Aplicar algún estilo básico al encabezado
                            // worksheet.Cells[FilaParaEscribir, i + 1].Style.Font.Bold = true;
                            // worksheet.Cells[FilaParaEscribir, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            // worksheet.Cells[FilaParaEscribir, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }
                        FilaParaEscribir++; // Siguiente fila para los datos
                    }
                    else
                    {
                        // Si la hoja ya tiene datos, encontrar la siguiente fila vacía
                        FilaParaEscribir = worksheet.Dimension.End.Row + 1;
                    }

                    // Escribir los datos de cada liquidación
                    foreach (var data in datosParaEscribir)
                    {
                        int col = 1;
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.Periodo;
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.Rut;
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.ApellidoPaterno;
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.ApellidoMaterno;
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.Nombres;

                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.SueldoBase);
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.CentroDeCosto;
                        WriteIntCell(worksheet, FilaParaEscribir, col++, data.DiasTrabajados);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Atraso);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Vacaciones);

                        worksheet.Cells[FilaParaEscribir, col++].Value = data.IsapreFonasa;
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.Plan;
                        worksheet.Cells[FilaParaEscribir, col++].Value = data.Afp;
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.PorcentajeAfp, "0.00\\%"); // Formato porcentaje

                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.SueldoMensual); // Pendiente celda origen
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Gratificacion);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.TotalImponible);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.CargaFamiliar);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.TotalNoImponible);

                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.MontoAfp);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Apv1);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Apv2);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.MontoSalud);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.SeguroCesantia);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.ImpuestoUnico);

                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.TotalDescuentos);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.TotalOtrosDescuentos);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.LiquidoAPagar);

                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Sis);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Mutual);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.AporteSeguroCesantiaEmpleador);
                        WriteDecimalCell(worksheet, FilaParaEscribir, col++, data.Tributable); // Pendiente celda origen

                        FilaParaEscribir++;
                    }

                    // Opcional: Autoajustar columnas para mejor visualización
                    // for(int i = 1; i <= Headers.Length; i++)
                    // {
                    //     worksheet.Column(i).AutoFit();
                    // }

                    package.Save();
                }
                Console.WriteLine($"Datos guardados exitosamente en: {rutaArchivoDestino} (Hoja: {anioSeleccionado})");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error al escribir el archivo de Excel: {ex.Message}");
                // Considerar si el archivo está abierto por otro programa
                if (ex is InvalidOperationException && ex.Message.Contains("locked"))
                {
                    Console.WriteLine("El archivo podría estar abierto por otra aplicación. Por favor, ciérrelo e intente de nuevo.");
                }
                return false;
            }
        }

        // Helpers para escribir celdas con tipos específicos y formato (opcional)
        private void WriteDecimalCell(ExcelWorksheet worksheet, int row, int col, decimal? value, string format = "#,##0.00")
        {
            if (value.HasValue)
            {
                worksheet.Cells[row, col].Value = value.Value;
                // worksheet.Cells[row, col].Style.Numberformat.Format = format; // Descomentar para aplicar formato específico
            }
            else
            {
                worksheet.Cells[row, col].Value = null; // O string.Empty si prefieres celdas vacías en lugar de nulas
            }
        }

        private void WriteIntCell(ExcelWorksheet worksheet, int row, int col, int? value, string format = "0")
        {
            if (value.HasValue)
            {
                worksheet.Cells[row, col].Value = value.Value;
                // worksheet.Cells[row, col].Style.Numberformat.Format = format; // Descomentar para aplicar formato específico
            }
            else
            {
                worksheet.Cells[row, col].Value = null;
            }
        }
    }
}
