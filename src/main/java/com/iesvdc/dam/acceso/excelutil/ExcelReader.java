package com.iesvdc.dam.acceso.excelutil;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.iesvdc.dam.acceso.modelo.FieldModel;
import com.iesvdc.dam.acceso.modelo.FieldType;
import com.iesvdc.dam.acceso.modelo.TableModel;
import com.iesvdc.dam.acceso.modelo.WorkbookModel;

public class ExcelReader {
    private Workbook wb;
    private WorkbookModel wbm;
    private final double EPSILON = 1e-10;

    public ExcelReader() {
    }

    /**
     * Devuelve un String indicando el tipo de dato de la celda.
     * Puede ser: Entero, Decimal, Texto, Booleano, Fecha, Vacía, Fórmula, Error
     */
    public FieldType getTipoDato(Cell cell) {
        if (cell == null) {
            return FieldType.UNKNOWN;
        }

        switch (cell.getCellType()) {
            case STRING:
                return FieldType.STRING;

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return FieldType.DATE;
                } else {
                    double valor = cell.getNumericCellValue();
                    if (Math.abs(valor - Math.floor(valor)) < EPSILON) {
                        return FieldType.INTEGER;
                    } else {
                        return FieldType.DECIMAL;
                    }
                }

            case BOOLEAN:
                return FieldType.BOOLEAN;

            default:
                return FieldType.UNKNOWN;
        }
    }

    public void loadWorkbook(String filename) {
        try (FileInputStream fis = new FileInputStream(filename)) {
            wb = new XSSFWorkbook(fis);
            wbm = new WorkbookModel();

            int nHojas = wb.getNumberOfSheets();
            for (int i = 0; i < nHojas; i++) {
                //Nombre de la tabla
                Sheet hojaActual = wb.getSheetAt(i);

                TableModel tabla = new TableModel(hojaActual.getSheetName());

                //Cojo el nombre de la primera fila
                Row primeraFila = hojaActual.getRow(0);

                //Cojo el tipo de dato de la segunda fila
                Row segundaFila = hojaActual.getRow(1);

                //Adquirimos el número de columnas
                int nCols = primeraFila.getLastCellNum();
                for (int j = 0; j < nCols; j++) {
                    FieldModel campo = 
                        new FieldModel (
                            primeraFila.getCell(j).getStringCellValue(),
                            getTipoDato(segundaFila.getCell(j))
                        );

                    tabla.addField(campo);
                    // Depuracion -> System.out.println("Añadiendo campo: " + campo.toString());

                }
                wbm.addTable(tabla);
            }

            System.out.println("Tablas:");
            System.out.println(wbm.toString());

        } catch (Exception e) {
            System.out.println("Imposible cargar el archivo Excel: " + e.getLocalizedMessage());
        }
    }
}
