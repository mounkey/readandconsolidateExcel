using System;
using System.Collections.Generic; // Para List<LiquidacionData>
using System.IO; // Para Path

namespace ReadAndConsolidateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("--- Iniciando el Programa de Consolidación de Liquidaciones ---");
            Console.WriteLine();

            // Configuración de Licencia EPPlus (solo una vez por AppDomain)
            // Descomentar si usas EPPlus 5 o superior y no tienes una licencia comercial global.
            // OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            // 1. Pedir ruta del archivo de origen
            Console.Write("Por favor, ingresa la ruta completa del archivo Excel de liquidación de origen: ");
            string? rutaArchivoOrigen = Console.ReadLine()?.Trim();

            if (string.IsNullOrWhiteSpace(rutaArchivoOrigen))
            {
                Console.WriteLine("Ruta de archivo de origen no válida. Saliendo del programa.");
                FinalizarPrograma();
                return;
            }

            if (!File.Exists(rutaArchivoOrigen))
            {
                Console.WriteLine($"Error: El archivo de origen no existe en: {rutaArchivoOrigen}. Saliendo del programa.");
                FinalizarPrograma();
                return;
            }

            // 2. Pedir el año correspondiente a la liquidación
            Console.Write("Por favor, ingresa el AÑO (ej: 2019, 2023) al que corresponde esta liquidación: ");
            string? anioInput = Console.ReadLine()?.Trim();
            if (string.IsNullOrWhiteSpace(anioInput) || !int.TryParse(anioInput, out int anioNumero) || anioInput.Length != 4)
            {
                Console.WriteLine("Año no válido. Debe ser un número de 4 dígitos. Saliendo del programa.");
                FinalizarPrograma();
                return;
            }
            string anioProcesamiento = anioInput; // Usaremos el string para el nombre de la hoja

            // 3. Leer la liquidación
            Console.WriteLine($"\nLeyendo datos de: {rutaArchivoOrigen}...");
            var lector = new ExcelDataReader();
            LiquidacionData? datosLeidos = lector.LeerLiquidacion(rutaArchivoOrigen);

            if (datosLeidos == null)
            {
                Console.WriteLine("No se pudieron leer los datos de la liquidación. Revisa los mensajes de error anteriores.");
                FinalizarPrograma();
                return;
            }

            Console.WriteLine("Lectura completada exitosamente.");
            // Console.WriteLine($"Datos leídos: {datosLeidos}"); // Para debugging

            // 4. Escribir en el consolidado
            // Definir ruta del archivo de destino (ej: en la misma carpeta del ejecutable o en Documentos)
            string nombreArchivoDestino = "Consolidado_Liquidaciones.xlsx";
            string rutaArchivoDestino = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nombreArchivoDestino);
            // string rutaArchivoDestino = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), nombreArchivoDestino); // Alternativa: Guardar en Documentos

            Console.WriteLine($"\nIntentando escribir datos en: {rutaArchivoDestino} (Hoja: {anioProcesamiento})");

            var escritor = new ExcelDataWriter();
            var listaDeDatos = new List<LiquidacionData> { datosLeidos }; // Creamos una lista aunque sea un solo elemento

            bool exitoEscritura = escritor.EscribirConsolidado(listaDeDatos, rutaArchivoDestino, anioProcesamiento);

            if (exitoEscritura)
            {
                Console.WriteLine("Proceso de consolidación completado exitosamente.");
            }
            else
            {
                Console.WriteLine("Falló el proceso de escritura en el archivo consolidado.");
            }

            FinalizarPrograma();
        }

        static void FinalizarPrograma()
        {
            Console.WriteLine("\n--- Programa finalizado. Presiona cualquier tecla para salir. ---");
            Console.ReadKey();
        }
    }
}
