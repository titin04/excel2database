# Proyecto excel2database

Programa ejemplo en Java que vuelca un archivo Excel a una base de datos MySQL.

Este programa genérico en java (proyecto Maven) es un ejercicio simple que vuelca un libro Excel (xlsx) a una base de datos (MySQL) y viceversa. El programa lee la configuración de la base de datos de un fichero "properties" de Java y luego, con apache POI, leo las hojas, el nombre de cada hoja será el nombre de las tablas, la primera fila de cada hoja será el nombre de los atributos de cada tabla (hoja) y para saber el tipo de dato, tendré que preguntar a la segunda fila qué tipo de dato tiene.

Procesamos el fichero Excel y creamos una estructura de datos con la información siguiente: La estructura principal es el libro, que contiene una lista de tablas y cada tabla contiene tuplas nombre del campo y tipo de dato.


Uso del programa:

```bash
excel2database -f fichero.xlsx -db agenda
```

## Instalando Java para que funcione

1. Instala **Java 17** o superior.  
   Ubuntu / Debian:
   ```bash
   sudo apt-get update
   sudo apt-get install openjdk-17-jdk
   java -version
   ```
2. Instala **Maven 3.9+**:
   ```bash
   sudo apt-get install maven
   mvn -v
   ```

3. Configura MySQL (puedes usar Docker con `stack-exelreader/docker-compose.yml`).

## Cómo crear este proyecto maven desde cero

Instalamos las dependencias Maven:

* apache poi
* apache poi ooxml
* mysql 

## El archivo de propiedades

Creamos el archivo **`config.properties`**:

```bash
user=root
password=s83n38DGB8d72
useUnicode=yes
useJDBCCompliantTimezoneShift=true
port=33307
database=agenda
host=localhost
driver=MySQL
outputFile=datos/salida.xlsx
inputFile=datos/entrada.xlsx
useSSL=false
serverTimezone=Europe/Madrid
allowPublicKeyRetrieval=true
```

En producción **jamás** debemos de usar estos parámetros:

* `useSSL=false`: No encripta la conexión.
* `allowPublicKeyRetrieval=true`: No comprueba el certificado (como el candado rojo del navegador)

## Detectando qué tipo de dato hay con Apache POI

Con **Apache POI** puedes inspeccionar el **tipo de dato almacenado en una celda de Excel (.xlsx)** y actuar según corresponda. Cuando trabajas con una celda (`Cell`), POI te permite preguntar su tipo con:

```java
cell.getCellType()
```

Esto devuelve un valor del enum `CellType`, que puede ser:

| Tipo (`CellType`) | Significado                                                             |
| ----------------- | ----------------------------------------------------------------------- |
| `NUMERIC`         | Número (entero o decimal, o incluso fecha/hora si el formato lo indica) |
| `STRING`          | Texto                                                                   |
| `BOOLEAN`         | Verdadero/Falso                                                         |
| `FORMULA`         | Celda con una fórmula                                                   |
| `BLANK`           | Celda vacía                                                             |
| `ERROR`           | Celda con error                                                         |


Excel almacena las **fechas y horas como números** (un número de días desde el 1/1/1900).
Para distinguirlas, Apache POI ofrece un método:

```java
DateUtil.isCellDateFormatted(cell)
```

Si devuelve `true`, el contenido es una **fecha o una hora**, y puedes obtenerla así:

```java
Date fecha = cell.getDateCellValue();
```

Si no, puedes tratarlo como número:

```java
double valor = cell.getNumericCellValue();
```

En Excel, todos los números se almacenan como double internamente. No existe distinción “formal” entre enteros y decimales dentro del archivo. Por eso, Apache POI siempre te devuelve NUMERIC y cell.getNumericCellValue() da un double. Como Excel no distingue formalmente entre ellos: ambos son `NUMERIC`, podemos comprobarlo fácilmente así:

```java
double valor = cell.getNumericCellValue();
if (valor == Math.floor(valor)) {
    System.out.println("Entero: " + (int) valor);
} else {
    System.out.println("Decimal: " + valor);
}
```

Un ejemplo de cómo hacer la detección sería:

```java

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;

public class ExcelUtils {

    private static final double EPSILON = 1e-10;

    /**
     * Devuelve un String indicando el tipo de dato de la celda.
     * Puede ser: Entero, Decimal, Texto, Booleano, Fecha, Vacía, Fórmula, Error
     */
    public static String getTipoDato(Cell cell) {
        if (cell == null) {
            return "Vacía";
        }

        switch (cell.getCellType()) {
            case STRING:
                return "Texto";

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return "Fecha";
                } else {
                    double valor = cell.getNumericCellValue();
                    if (Math.abs(valor - Math.floor(valor)) < EPSILON) {
                        return "Entero";
                    } else {
                        return "Decimal";
                    }
                }

            case BOOLEAN:
                return "Booleano";

            case FORMULA:
                // Puedes decidir si quieres evaluar la fórmula o solo indicar que es fórmula
                return "Fórmula";

            case BLANK:
                return "Vacía";

            case ERROR:
                return "Error";

            default:
                return "Desconocido";
        }
    }
}

```

## Solución de problemas

- `mvn: command not found` → instala Maven.
- `No se pudo obtener la conexión` → revisa host, puerto, usuario y que MySQL esté activo.
- Tipos incorrectos → verifica la fila 1 del Excel.
- Excel vacío al exportar → comprueba el esquema (`database`) y permisos.

## Datos de ejemplo

- `datos/test.xlsx`: 19 registros de prueba.
- `datos/personas.xlsx`: libro adicional.
- `datos/salida.xlsx`: se crea al exportar.

## Scripts y utilidades

- `stack-exelreader/docker-compose.yml`: levanta MySQL con datos de ejemplo.
- `docs/`: documentación complementaria.

## Ejemplos de ejecución

### Importar (acción `load`)
```
El nombre del archivo es: datos/test.xlsx
La acción es: load
Generando tablas...
CREATE TABLE personas(`nombre` VARCHAR(255), ...)
Importación completada con éxito.
```

### Exportar (acción `save`)
```
El nombre del archivo es: datos/test.xlsx
La acción es: save
Exportación completada. Archivo generado en: datos/salida.xlsx
```

Si aparece `No se pudo obtener la conexión a la base de datos`, revisa credenciales y que MySQL esté corriendo.

## Estructura del proyecto

```
datos/                     # Excels de entrada/salida
docs/                      # Documentación y utilidades
src/main/java/
  com/iesvdc/dam/acceso/
    ExcelDatabaseApp.java  # Clase principal (main)
    conexion/              # Config y conexión JDBC
    databaseutil/          # Exportación MySQL -> Excel
    excelutil/             # Importación Excel -> MySQL
    modelo/                # POJOs del libro Excel
```

## Comprobaciones rápidas

1. **Test de compilación**  
   ```
   mvn -q -DskipTests compile
   ```
2. **Conexión a MySQL**  
   Asegúrate de que `config.properties` apunta al servidor correcto. Puedes ejecutar solo la exportación (`action=save`) para comprobar que la conexión funciona.
3. **Libro de salida**  
   Tras exportar, abre `datos/salida.xlsx` y verifica:
   - Fila 0: nombres de columnas.
   - Filas siguientes: datos reales sin sobrescribir la fila 1.

## Referencias adicionales

- `docs/Tema02TareaExcel2Database.pdf`: enunciado completo con rúbrica.
- `docs/makebook.sh`: script para generar el PDF en Linux usando `pandoc`.

## Apéndice: Repaso de SQL

Para crear una base de datos en MySQL hacemos:

```sql
create database `agenda`;
```

Para borrar una base de datos en MySQL hacemos:

```sql 
drop database `agenda`;
```


```sql
CREATE DATABASE `agenda` COLLATE 'utf16_spanish_ci';
```

Crear una tabla de personas:

```sql
CREATE TABLE `personas` (
  `nombre` varchar(100) NOT NULL,
  `apellidos` varchar(300) NOT NULL,
  `email` varchar(100),
  `telefono` varchar(12),
  `genero` enum('FEMENINO','MASCULINO','NEUTRO','OTRO') NOT NULL,
  `id` int NOT NULL AUTO_INCREMENT PRIMARY KEY
) ENGINE='InnoDB';



INSERT INTO `personas` (
    `nombre`, `apellidos`, 
    `email`, `telefono`, 
    `genero`)
VALUES (
    'Juan', 'Sin Miedo', 
    'juan@sinmiedo.com', '+34555123456', 
    'OTRO');

SELECT * FROM `personas` LIMIT 50;

SELECT * FROM `personas` LIMIT 5 OFFSET 10;
```
