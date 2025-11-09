package com.iesvdc.dam.acceso;

import java.util.Properties;

import com.iesvdc.dam.acceso.conexion.Config;
import com.iesvdc.dam.acceso.databaseutil.DatabaseReader;
import com.iesvdc.dam.acceso.excelutil.ExcelReader;

/**
 * Aplicación principal que permite importar datos desde un Excel a MySQL
 * o exportar la base de datos a un nuevo fichero Excel.
 */
public class ExcelDatabaseApp {
    public static void main(String[] args) {
        Properties props = Config.getProperties("config.properties"); // Cargar configuración desde el fichero

        System.out.println("El nombre del archivo es: " + props.getProperty("file"));
        System.out.println("La acción es: " + props.getProperty("action"));
        String action = props.getProperty("action", "load").toLowerCase(); // Acción por defecto es 'load'
        switch (action) {
            // Importar desde Excel a la base de datos
            case "load":
            case "import":
                importarExcel(props);
                break;

            // Exportar desde la base de datos a Excel
            case "save":
            case "export":
                exportarBaseDatos(props);
                break;

            // Acción no reconocida
            default:
                System.out.println("Acción no reconocida: " + action);
                System.out.println("Usa 'load' para importar Excel o 'save' para exportar la base de datos.");
        }
    }

    /**
     * Importa datos desde un fichero Excel a la base de datos.
     * 
     * @param props
     */
    private static void importarExcel(Properties props) {
        String inputFile = props.getProperty("inputFile", props.getProperty("file")); // Obtener ruta del fichero de entrada
        if (inputFile == null) {
            System.err.println("No se ha indicado el fichero de entrada (inputFile).");
            return;
        }

        ExcelReader reader = new ExcelReader(); // Crear instancia del lector de Excel
        reader.loadWorkbook(inputFile); // Cargar el libro de Excel
        try {
            reader.saveToDatabase(); // Guardar datos en la base de datos
            System.out.println("Importación completada con éxito.");
        } catch (RuntimeException ex) {
            System.err.println("Error al importar el Excel a la base de datos: " + ex.getMessage());
        }
    }

    /**
     * Exporta los datos de la base de datos a un fichero Excel.
     * 
     * @param props
     */
    private static void exportarBaseDatos(Properties props) {
        String outputFile = props.getProperty("outputFile", "datos/salida.xlsx"); // Obtener ruta del fichero de salida
        DatabaseReader exporter = new DatabaseReader(); // Crear instancia del exportador de base de datos
        try {
            exporter.export(outputFile); // Exportar datos a Excel
            System.out.println("Exportación completada. Archivo generado en: " + outputFile);
        } catch (Exception e) {
            System.err.println("Error al exportar la base de datos a Excel: " + e.getMessage());
        }
    }
}
