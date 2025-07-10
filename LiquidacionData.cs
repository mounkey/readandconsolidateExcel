using System;

namespace ReadAndConsolidateExcel
{
    public class LiquidacionData
    {
        // Propiedades basadas en la cabecera del archivo de destino
        // El A�O se manejar� por el nombre de la hoja en el archivo de destino
        // y se pedir� al usuario, por lo que no es una propiedad aqu�.
        // PERIODO almacenar� el MES.

        public string? Periodo { get; set; } // Mes (ej: "MARZO")
        public string? Rut { get; set; }
        public string? ApellidoPaterno { get; set; }
        public string? ApellidoMaterno { get; set; }
        public string? Nombres { get; set; }
        public decimal? SueldoBase { get; set; }
        public string? CentroDeCosto { get; set; }
        public int? DiasTrabajados { get; set; }
        public decimal? Atraso { get; set; } // Asumiendo que es un valor monetario, si es tiempo, cambiar tipo
        public decimal? Vacaciones { get; set; } // Asumiendo valor monetario o d�as, ajustar tipo si es necesario
        public string? IsapreFonasa { get; set; } // Nombre de la instituci�n
        public string? Plan { get; set; }
        public string? Afp { get; set; } // Nombre de la instituci�n
        public decimal? PorcentajeAfp { get; set; } // Ej: 10.77 para 10.77%
        public decimal? SueldoMensual { get; set; } // Pendiente de confirmaci�n de celda origen
        public decimal? Gratificacion { get; set; }
        public decimal? TotalImponible { get; set; }
        public decimal? CargaFamiliar { get; set; }
        public decimal? TotalNoImponible { get; set; }
        public decimal? MontoAfp { get; set; } // El descuento de AFP
        public decimal? Apv1 { get; set; }
        public decimal? Apv2 { get; set; }
        public decimal? MontoSalud { get; set; } // El descuento de Salud
        public decimal? SeguroCesantia { get; set; } // El descuento de Seguro Cesant�a
        public decimal? ImpuestoUnico { get; set; }
        public decimal? TotalDescuentos { get; set; }
        public decimal? TotalOtrosDescuentos { get; set; } // Se dejar� en blanco/nulo por ahora
        public decimal? LiquidoAPagar { get; set; }
        public decimal? Sis { get; set; } // Seguro Invalidez y Sobrevivencia (aporte empleador o descuento?)
        public decimal? Mutual { get; set; }
        public decimal? AporteSeguroCesantiaEmpleador { get; set; } // Si es diferente al descuento del trabajador
        public decimal? Tributable { get; set; } // Pendiente de confirmaci�n de celda origen

        // Constructor por si es �til
        public LiquidacionData() { }

        // Podr�amos a�adir un m�todo ToString() para debugging f�cil
        public override string ToString()
        {
            return $"{Periodo} - {Rut} - {ApellidoPaterno} {Nombres} - L�quido: {LiquidoAPagar}";
        }
    }
}
