# RA2: “Desarrolla aplicaciones que gestionan información almacenada en bases de datos relacionales identificando y utilizando mecanismos de conexión”*

## Proyecto: *Excel2Database*

### Descripción general

El proyecto **Excel2Database** tiene como objetivo desarrollar una aplicación **Java (Maven)** capaz de **importar y exportar información** entre un archivo **Excel (.xlsx)** y una **base de datos MySQL**.

Este ejercicio permitirá al alumnado **consolidar los fundamentos del acceso a datos relacionales** mediante JDBC, el uso de ficheros, la manipulación de estructuras de datos y la configuración de proyectos Maven con dependencias externas.


### Objetivos de aprendizaje

* Comprender el funcionamiento del **conector JDBC** y el proceso de conexión a MySQL.
* Aplicar el acceso a datos relacionales mediante **sentencias SQL** desde Java.
* Practicar la **lectura y escritura de ficheros Excel** mediante la librería **Apache POI**.
* Diseñar un programa estructurado y modular que cumpla un objetivo realista y completo.
* Configurar la aplicación a través de un fichero **`.properties`** externo.
* Integrar herramientas de gestión de dependencias (Maven) y documentación del código.


### Descripción funcional

El programa debe ejecutarse desde línea de comandos con esta sintaxis:

```
excel2database -f fichero.xlsx -db nombreBD
```

#### Fases del proyecto:

1. **Lectura de configuración**

   * El programa cargará los parámetros de conexión desde un fichero `config.properties`:

     ```
     host=localhost
     port=3306
     database=agenda
     user=root
     password=****
     ```
2. **Análisis del libro Excel**

   * Cada **hoja** del libro representa una **tabla** de la base de datos.
   * La **primera fila** contiene los nombres de los campos.
   * La **segunda fila** define el **tipo de dato** (por ejemplo: `INT`, `VARCHAR(50)`, `DATE`).
   * Las filas siguientes son los **registros**.
3. **Creación de la base de datos y tablas**

   * Se generarán automáticamente las sentencias SQL `CREATE TABLE` según la estructura detectada.
4. **Inserción de los datos**

   * Se leerán los registros del Excel e insertarán en la base de datos con sentencias `INSERT`.
5. **Exportación inversa**

   * El programa permitirá la operación inversa: exportar una base de datos MySQL a un nuevo fichero Excel.


### Requisitos técnicos

* **Lenguaje:** Java 17 o superior
* **Gestor de dependencias:** Maven
* **Dependencias mínimas:**

  ```xml
  <dependencies>
      <dependency>
          <groupId>org.apache.poi</groupId>
          <artifactId>poi-ooxml</artifactId>
          <version>5.2.5</version>
      </dependency>
      <dependency>
          <groupId>mysql</groupId>
          <artifactId>mysql-connector-j</artifactId>
          <version>8.4.0</version>
      </dependency>
  </dependencies>
  ```
* El código debe incluir **tratamiento de excepciones**, **mensajes informativos** y **comentarios de documentación** (`/** ... */`).


## Rúbrica de evaluación (RA2)

| **Criterio de evaluación (según PD)**                          | **Indicadores observables**                                                             | **Excelente (9-10)**                                                               | **Notable (7-8)**                                    | **Aprobado (5-6)**                                     | **Insuficiente (<5)**           | **Peso** |
| -------------------------------------------------------------- | --------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------- | ---------------------------------------------------- | ------------------------------------------------------ | ------------------------------- | -------- |
| **a)** Valora las ventajas e inconvenientes de los conectores. | Explica en memoria el uso del conector MySQL y las razones de su elección.              | Justifica correctamente el uso de JDBC y las ventajas frente a otras alternativas. | Describe adecuadamente el uso de JDBC, con ejemplos. | Lo usa sin justificarlo.                               | No explica ni comprende su uso. | 10%      |
| **b)** Utiliza gestores embebidos o independientes.            | Prueba la conexión con MySQL desde código y demuestra comprensión.                      | Configura correctamente la conexión y demuestra autonomía.                         | Configura con ayuda o correcciones.                  | Configura parcialmente.                                | No consigue conectar.           | 10%      |
| **c)** Usa el conector idóneo en la aplicación.                | Incluye correctamente el driver JDBC en `pom.xml`.                                      | Correcta configuración y documentación.                                            | Correcta pero sin documentar.                        | Usa dependencias redundantes o incompletas.            | No funciona el conector.        | 10%      |
| **d)** Establece la conexión.                                  | Ejecuta correctamente `DriverManager.getConnection()` con parámetros del `.properties`. | Conexión estable y gestionada con cierre correcto (`try-with-resources`).          | Conexión correcta, cierre manual.                    | Conexión funciona pero sin gestionar bien excepciones. | Error de conexión o parámetros. | 5%       |
| **e)** Define la estructura de la base de datos.               | Crea las tablas desde el Excel analizado.                                               | Genera correctamente el SQL dinámico `CREATE TABLE`.                               | Genera las tablas con pequeñas imprecisiones.        | Las crea parcialmente.                                 | No crea tablas válidas.         | 10%      |
| **f)** Desarrolla aplicaciones que modifican el contenido.     | Inserta registros correctamente.                                                        | Inserta todos los datos sin error y controla duplicados.                           | Inserta correctamente con leves errores.             | Inserta parcialmente.                                  | No inserta datos.               | 15%      |
| **g)** Define objetos para almacenar resultados.               | Crea clases para representar tablas y campos.                                           | Estructura clara con POJOs y uso de colecciones.                                   | Clases correctas sin estructura óptima.              | Estructura funcional mínima.                           | No modela los datos.            | 10%      |
| **h)** Desarrolla aplicaciones que efectúan consultas.         | Implementa lectura y exportación de datos a Excel.                                      | Exporta correctamente con Apache POI.                                              | Exporta con pequeños fallos.                         | Exporta parcialmente.                                  | No exporta.                     | 15%      |
| **i)** Elimina objetos al finalizar su función.                | Libera recursos (streams, conexiones).                                                  | Cierra todos los recursos correctamente.                                           | Cierra parcialmente.                                 | Cierre manual incompleto.                              | No libera recursos.             | 5%       |
| **j)** Gestiona transacciones.                                 | Aplica `commit` y `rollback` correctamente.                                             | Controla transacciones correctamente con feedback al usuario.                      | Usa transacciones básicas.                           | Sin gestión explícita.                                 | Genera inconsistencias.         | 10%      |

**Total:** 100%
