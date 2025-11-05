package com.iesvdc.dam.acceso.modelo;
import java.util.Objects;

/**
 * El modelo que almacena información de un campo y sus propiedades.
 */

public class FieldModel {
    private final String name;
    private final FieldType type;



    public FieldModel() {
        this.name = "";
        this.type = FieldType.UNKNOWN;
    }

    public FieldModel(String name) {
        this.name = name;
        this.type = FieldType.UNKNOWN;
    }
    
    public FieldModel(String name, FieldType type) {
        this.name = name;
        this.type = type;
    }

    public String getName() {
        return this.name;
    }


    public FieldType getType() {
        return this.type;
    }


    @Override
    public boolean equals(Object o) {
        if (o == this)
            return true;
        if (!(o instanceof FieldModel)) {
            return false;
        }
        FieldModel fieldModel = (FieldModel) o;
        return Objects.equals(name, fieldModel.name) && Objects.equals(type, fieldModel.type);
    }


    @Override
    public String toString() {
        return "{" +
            " name='" + getName() + "'" +
            ", type='" + getType() + "'" +
            "}";
    }
    
}
