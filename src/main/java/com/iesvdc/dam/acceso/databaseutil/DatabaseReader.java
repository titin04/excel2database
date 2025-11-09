package com.iesvdc.dam.acceso.databaseutil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.iesvdc.dam.acceso.conexion.Conexion;

/**
 * Clase de ayuda que lee el contenido de la base de datos y lo
 * transforma en un {@link org.apache.poi.ss.usermodel.Workbook}.
 *
 * Cada tabla se exporta como una hoja:
 * <ul>
 *  <li>Fila 0: nombres de columnas.</li>
 *  <li>Fila 1: tipos de datos según la base de datos.</li>
 *  <li>Filas siguientes: registros existentes.</li>
 * </ul>
 */
public class DatabaseReader {

    /**
     * Lee todas las tablas visibles en la base de datos actual y las vuelca
     * en un libro de Excel.
     *
     * @return workbook con la representación de las tablas.
     * @throws SQLException si ocurre cualquier error al consultar la base de datos.
     */
    public Workbook readDatabase() throws SQLException {
        try (Connection connection = Conexion.getConnection()) {
            if (connection == null) {
                throw new SQLException("No se pudo obtener la conexión a la base de datos.");
            }

            Workbook workbook = new XSSFWorkbook();
            DatabaseMetaData metaData = connection.getMetaData(); // Obtener metadatos de la base de datos
            String catalog = connection.getCatalog(); // Obtener el catálogo actual

            try (ResultSet tables = metaData.getTables(catalog, null, "%", new String[] { "TABLE" })) { // Obtener todas las tablas
                while (tables.next()) { // Iterar sobre cada tabla
                    String tableName = tables.getString("TABLE_NAME"); // Nombre de la tabla
                    exportTable(workbook, connection, tableName); // Exportar la tabla al workbook
                }
            }

            return workbook;
        }
    }

    /**
     * Exporta todas las tablas de la base de datos a un fichero Excel.
     *
     * @param outputPath ruta del fichero .xlsx de salida.
     * @throws SQLException si ocurre un error al consultar la base de datos.
     * @throws IOException si no se puede crear o escribir el fichero.
     */
    public void export(String outputPath) throws SQLException, IOException {
        Workbook workbook = readDatabase();

        File outputFile = new File(outputPath); // Crear el archivo de salida
        File parent = outputFile.getParentFile(); // Obtener el directorio padre
        if (parent != null && !parent.exists()) {
            parent.mkdirs(); // Crear directorios padre si no existen
        }

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    /**
     * Exporta una tabla concreta a una hoja dentro del {@link Workbook} indicado.
     * Se crea una hoja con el nombre de la tabla, encabezados con los nombres de columna,
     * una segunda fila con los tipos SQL y las filas posteriores con los registros.
     *
     * @param workbook   libro donde se añadirá la hoja.
     * @param connection conexión activa a la base de datos.
     * @param tableName  nombre de la tabla que se va a exportar.
     * @throws SQLException si la consulta falla.
     */
    private void exportTable(Workbook workbook, Connection connection, String tableName) throws SQLException {
        Sheet sheet = workbook.createSheet(tableName); // Crear una nueva hoja para la tabla

        try (Statement statement = connection.createStatement(); // Crear una declaración SQL
            ResultSet rs = statement.executeQuery("SELECT * FROM `" + tableName + "`")) {

            ResultSetMetaData rsMeta = rs.getMetaData(); // Obtener metadatos del conjunto de resultados
            int columnCount = rsMeta.getColumnCount(); // Obtener el número de columnas

            Row headerRow = sheet.createRow(0); // Crear la fila de encabezado

            for (int i = 1; i <= columnCount; i++) { // Iterar sobre cada columna
                Cell headerCell = headerRow.createCell(i - 1); 
                headerCell.setCellValue(rsMeta.getColumnLabel(i)); 
            }

            int rowIndex = 1; // Índice para las filas de datos
            while (rs.next()) { // Iterar sobre cada fila del conjunto de resultados
                Row dataRow = sheet.createRow(rowIndex++); 
                for (int i = 1; i <= columnCount; i++) {
                    Cell cell = dataRow.createCell(i - 1);
                    Object value = rs.getObject(i);
                    setCellValue(cell, value); // Establecer el valor de la celda
                }
            }
        }
    }

    /**
     * Escribe un valor Java en una celda Excel utilizando un formato legible.
     *
     * @param cell  celda a rellenar.
     * @param value valor obtenido desde la base de datos.
     */
    private void setCellValue(Cell cell, Object value) {
        if (value == null) {
            return;
        }

        // Determinar el tipo de dato y establecer el valor en la celda
        if (value instanceof Number number) {
            cell.setCellValue(number.doubleValue());
        } else if (value instanceof java.sql.Date date) {
            cell.setCellValue(date.toLocalDate().toString());
        } else if (value instanceof java.sql.Timestamp timestamp) {
            cell.setCellValue(timestamp.toLocalDateTime().toString());
        } else if (value instanceof Boolean bool) {
            cell.setCellValue(bool);
        } else {
            cell.setCellValue(value.toString()); // Tratar como cadena por defecto
        }
    }
}