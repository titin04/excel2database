package com.iesvdc.dam.acceso.modelo;

/**
 * Tipos de datos posibles para los campos
 */
public enum FieldType {
    INTEGER,
    FLOAT,
    VARCHAR,
    DATE,
    BOOLEAN,
    UNKNOWN;

    /**
     * Determina si el tipo representa un valor numérico.
     * @return true si es INTEGER o FLOAT.
     */
    public boolean isNumeric() {
        return this == INTEGER || this == FLOAT;
    }

    /**
     * Convierte el tipo en un tipo SQL estándar.
     * @return Cadena con el tipo compatible con MySQL.
     */
    public String toSqlType() {
        return switch (this) {
            case INTEGER -> "INT";
            case FLOAT -> "DOUBLE";
            case VARCHAR -> "VARCHAR(255)";
            case DATE -> "DATE";
            case BOOLEAN -> "BOOLEAN";
            default -> "VARCHAR(255)";
        };
    }
}