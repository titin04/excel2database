package com.iesvdc.dam.acceso.modelo;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * El modelo que almacena información de una tabla y su lista de campos.
 */
public class TableModel {
    private final String name;
    private final List<FieldModel> fields = new ArrayList<>();
    private final List<List<Object>> rows = new ArrayList<>();



    public TableModel() {
        this.name = "";
    }

    public TableModel(String name) {
        this.name = name;
    }

    public String getName() {
        return this.name;
    }

    public boolean addField(FieldModel fm) {
        return fields.add(fm);
    }

    public List<FieldModel> getFields() {
        return this.fields;
    }

    public boolean addRow(List<Object> row) {
        return rows.add(row);
    }

    public List<List<Object>> getRows() {
        return this.rows;
    }

    @Override
    public boolean equals(Object o) {
        if (o == this)
            return true;
        if (!(o instanceof TableModel)) {
            return false;
        }
        TableModel tableModel = (TableModel) o;
        return Objects.equals(name, tableModel.name)
            && Objects.equals(fields, tableModel.fields)
            && Objects.equals(rows, tableModel.rows);
    }


    @Override
    public String toString() {
        return "{" +
            " name='" + getName() + "'" +
            ", fields='" + getFields() + "'" +
            ", rows='" + getRows() + "'" +
            "}";
    }
    
}
