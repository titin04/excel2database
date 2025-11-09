package com.iesvdc.dam.acceso.excelutil;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.iesvdc.dam.acceso.conexion.Conexion;
import com.iesvdc.dam.acceso.modelo.FieldModel;
import com.iesvdc.dam.acceso.modelo.FieldType;
import com.iesvdc.dam.acceso.modelo.TableModel;
import com.iesvdc.dam.acceso.modelo.WorkbookModel;

/**
 * Clase auxiliar que interpreta un libro Excel y permite:
 * <ul>
 *   <li>Construir un modelo intermedio ({@link WorkbookModel}) con tablas, campos y filas.</li>
 *   <li>Generar y ejecutar el DDL necesario para recrear esas tablas en MySQL.</li>
 *   <li>Insertar los datos leídos en las tablas recién creadas.</li>
 * </ul>
 * Se apoya en Apache POI para leer el Excel y en JDBC para hablar con la base de datos.
 */
public class ExcelReader {
    /** Libro de POI con el contenido Excel original. */
    private Workbook wb;
    /** Modelo intermedio que describe todas las tablas, campos y filas. */
    private WorkbookModel wbm;
    /** Conexión JDBC usada para crear tablas e insertar filas. */
    private Connection conexion;
    /** Margen de error para decidir si un número es entero o decimal. */
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

        CellType tipoCelda = cell.getCellType();

        switch (tipoCelda) {
            case STRING:
                return FieldType.VARCHAR;

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return FieldType.DATE;
                } else {
                    double valor = cell.getNumericCellValue();
                    if (Math.abs(valor - Math.floor(valor)) < EPSILON) {
                        return FieldType.INTEGER;
                    } else {
                        return FieldType.FLOAT;
                    }
                }

            case BOOLEAN:
                return FieldType.BOOLEAN;

            default:
                return FieldType.UNKNOWN;
        }
    }

    /**
     * Carga un fichero Excel y convierte cada hoja en una {@link TableModel}.
     *
     * <p>Convenciones:</p>
     * <ul>
     *   <li>Fila 0: nombres de columna.</li>
     *   <li>Fila 1: ejemplos para detectar el tipo de dato.</li>
     *   <li>Filar 2 en adelante: registros a importar.</li>
     * </ul>
     *
     * @param filename ruta al fichero Excel (.xlsx).
     */
    public void loadWorkbook(String filename) {
        try (FileInputStream fis = new FileInputStream(filename)) {
            // Abrimos el libro y creamos el modelo raíz.
            wb = new XSSFWorkbook(fis);
            wbm = new WorkbookModel();

            int nHojas = wb.getNumberOfSheets();
            for (int i = 0; i < nHojas; i++) {
                Sheet hojaActual = wb.getSheetAt(i);
                // El nombre de la hoja pasa a ser el nombre de la tabla.
                TableModel tabla = new TableModel(hojaActual.getSheetName());

                // Fila 0 -> nombres de los campos.
                Row primeraFila = hojaActual.getRow(0);
                // Fila 1 -> valores de ejemplo para inferir el tipo.
                Row segundaFila = hojaActual.getRow(1);

                int nCols = primeraFila.getLastCellNum();
                List<FieldModel> campos = new ArrayList<>();

                // Recorremos las columnas para crear los FieldModel (nombre + tipo).
                for (int j = 0; j < nCols; j++) {
                    FieldModel campo = new FieldModel(
                        primeraFila.getCell(j).getStringCellValue(),
                        getTipoDato(segundaFila.getCell(j))
                    );

                    tabla.addField(campo);
                    campos.add(campo);
                }

                // A partir de la fila 2 están los datos reales.
                int filaActual = 2;
                Row filaDatos;
                while ((filaDatos = hojaActual.getRow(filaActual)) != null) {
                    List<Object> valores = new ArrayList<>();
                    boolean filaVacia = true;

                    for (int j = 0; j < nCols; j++) {
                        FieldModel campo = campos.get(j);
                        Cell celda = filaDatos.getCell(j);
                        Object valor = readCellValue(celda, campo.getType());
                        valores.add(valor);

                        if (valor != null) {
                            filaVacia = false;
                        }
                    }

                    // Solo añadimos filas que tengan al menos un valor.
                    if (!filaVacia) {
                        tabla.addRow(valores);
                    }

                    filaActual++;
                }

                wbm.addTable(tabla);
            }

        } catch (Exception e) {
            System.out.println("Imposible cargar el archivo Excel: " + e.getLocalizedMessage());
        }
    }

    /**
     * Genera y ejecuta el DDL necesario para crear las tablas definidas
     * en el {@link WorkbookModel} cargado previamente.
     *
     * @return true si todas las tablas se crearon correctamente.
     */
    public boolean executeDDL() {
        if (conexion == null) {
            conexion = Conexion.getConnection();
        }
        if (conexion == null) {
            throw new RuntimeException("No se pudo obtener conexión para crear las tablas.");
        }

        boolean resultado = true;

        for (TableModel tableModel : wbm.getTables()) {
            // Generamos el SQL CREATE TABLE usando los nombres y tipos del modelo.
            StringBuilder sqlSB = new StringBuilder();
            sqlSB.append("CREATE TABLE ");
            sqlSB.append(tableModel.getName());
            sqlSB.append("(");

            int nCampos = tableModel.getFields().size();
            for (FieldModel fieldModel : tableModel.getFields()) {
                nCampos--;
                sqlSB.append("`");
                sqlSB.append(fieldModel.getName());
                sqlSB.append("` ");
                sqlSB.append(fieldModel.getType().toSqlType());
                if (nCampos > 0) {
                    sqlSB.append(", ");
                }
            }
            sqlSB.append(");");

            try (Statement stmt = conexion.createStatement()) {
                // Eliminamos la tabla si ya existía y creamos la nueva estructura.
                stmt.execute("DROP TABLE IF EXISTS `" + tableModel.getName() + "`");
                stmt.executeUpdate(sqlSB.toString());
            } catch (SQLException e) {
                resultado = false;
                throw new RuntimeException("Error al crear la tabla " + tableModel.getName(), e);
            }
        }

        return resultado;
    }

        /**
     * Importa el libro Excel cargado a la base de datos.
     * <p>Pasos:</p>
     * <ol>
     *   <li>Validar que existe un {@link WorkbookModel} cargado.</li>
     *   <li>Abrir una conexión JDBC y desactivar el auto-commit para agrupar todas las operaciones.</li>
     *   <li>Recrear las tablas detectadas en el Excel con {@link #executeDDL()}.</li>
     *   <li>Insertar todas las filas de cada tabla con {@link #insertarTabla(TableModel)}.</li>
     *   <li>Confirmar los cambios; si algo falla, revertir con <code>rollback()</code>.</li>
     * </ol>
     */
    public void saveToDatabase() {
        if (wbm == null) {
            throw new IllegalStateException("Debe cargar primero un libro Excel antes de guardar en la base de datos.");
        }

        // try-with-resources asegura que la conexión se cierre aunque se produzca un error.
        try (Connection conn = Conexion.getConnection()) {
            if (conn == null) {
                throw new RuntimeException("No se pudo establecer la conexión con la base de datos.");
            }

            // Guardamos la referencia para que otros métodos (insertarTabla) puedan usarla.
            this.conexion = conn;
            // Desactivamos el auto-commit: todas las sentencias formarán parte de la misma transacción.
            conn.setAutoCommit(false);

            // 1) Crear tablas según el contenido del Excel.
            executeDDL();

            // 2) Insertar los datos de cada tabla.
            for (TableModel table : wbm.getTables()) {
                insertarTabla(table);
            }

            // 3) Confirmar la transacción: todas las operaciones quedan guardadas definitivamente.
            conn.commit();
        } catch (Exception e) {
            // Si algo falla, intentamos revertir los cambios realizados en esta transacción.
            if (conexion != null) {
                try {
                    conexion.rollback();
                } catch (SQLException ignore) { }
            }
            throw new RuntimeException("Error al volcar los datos del Excel a la base de datos.", e);
        } 

        // Limpiamos la referencia para evitar reutilizar una conexión cerrada.
        conexion = null;
    }

    /**
     * Inserta todas las filas de un {@link TableModel} en la base de datos.
     *
     * @param table tabla con la información procedente del Excel.
     */
    private void insertarTabla(TableModel table) throws SQLException {
        // No hay filas -> no hacemos nada.
        if (table.getRows().isEmpty()) {
            return;
        }

        // Montamos la sentencia INSERT con las columnas de la tabla.
        StringBuilder sql = new StringBuilder();
        sql.append("INSERT INTO `").append(table.getName()).append("` (");

        int nCampos = table.getFields().size();
        for (int i = 0; i < nCampos; i++) {
            FieldModel field = table.getFields().get(i);
            sql.append("`").append(field.getName()).append("`");
            if (i < nCampos - 1) {
                sql.append(", ");
            }
        }
        sql.append(") VALUES (");
        for (int i = 0; i < nCampos; i++) {
            sql.append("?");
            if (i < nCampos - 1) {
                sql.append(", ");
            }
        }
        sql.append(")");

        // Preparamos la sentencia y vamos añadiendo cada fila del Excel como lote.
        try (PreparedStatement ps = conexion.prepareStatement(sql.toString())) {
            for (List<Object> row : table.getRows()) {
                for (int i = 0; i < nCampos; i++) {
                    FieldType type = table.getFields().get(i).getType();
                    Object value = i < row.size() ? row.get(i) : null;
                    setPreparedValue(ps, i + 1, type, value);
                }
                ps.addBatch();
            }
            // Ejecutamos todas las filas de golpe para mayor eficiencia.
            ps.executeBatch();
        }
    }

    /**
     * Coloca un valor en la posición indicada del {@link PreparedStatement}
     * usando el método apropiado según el tipo detectado.
     */
    private void setPreparedValue(PreparedStatement ps, int index, FieldType type, Object value) throws SQLException {
        if (value == null) {
            ps.setNull(index, Types.NULL);
            return;
        }

        // Asignamos el valor según su tipo.
        switch (type) {
            case INTEGER:
                ps.setLong(index, ((Number) value).longValue());
                break;
            case FLOAT:
                ps.setDouble(index, ((Number) value).doubleValue());
                break;
            case DATE:
                ps.setDate(index, (java.sql.Date) value);
                break;
            case BOOLEAN:
                ps.setBoolean(index, (Boolean) value);
                break;
            case VARCHAR:
            case UNKNOWN:
                ps.setString(index, value.toString());
                break;
        }
    }

    /**
     * Traduce el contenido de una celda a un tipo Java compatible con JDBC,
     * respetando el tipo deducido previamente.
     */
    private Object readCellValue(Cell celda, FieldType tipo) {
        if (celda == null) {
            return null;
        }

        CellType cellType = celda.getCellType();

        switch (tipo) {
            case INTEGER:
                if (cellType == CellType.NUMERIC) {
                    return Math.round(celda.getNumericCellValue());
                }
                break;
            case FLOAT:
                if (cellType == CellType.NUMERIC) {
                    return celda.getNumericCellValue();
                }
                break;
            case DATE:
                if (cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(celda)) {
                    return new java.sql.Date(celda.getDateCellValue().getTime());
                }
                break;

            case BOOLEAN:
                if (cellType == CellType.BOOLEAN) {
                    return celda.getBooleanCellValue();
                }
                break;

            case VARCHAR:
                if (cellType == CellType.STRING) {
                    String texto = celda.getStringCellValue();
                    return texto.isEmpty() ? null : texto;
                }
                return celda.toString();

            case UNKNOWN:
                return celda.toString();
            }

        return null;
    }
}
