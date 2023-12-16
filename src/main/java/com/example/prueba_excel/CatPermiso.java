package com.example.prueba_excel;

public class CatPermiso {
    private int idPermiso;
    private String descripcion;
    private int activo;

    public CatPermiso() {
        super();
    }

    public CatPermiso(int idPermiso, String descripcion, int activo) {
        super();
        this.idPermiso = idPermiso;
        this.descripcion = descripcion;
        this.activo = activo;
    }

    public CatPermiso(PermisoDTO permisoDTO) {
        this.idPermiso = permisoDTO.idPermiso();
        descripcion = permisoDTO.descripcion();
        activo = permisoDTO.estatus();
    }

    public int getIdPermiso() {
        return this.idPermiso;
    }

    public void setIdPermiso(int idPermiso) {
        this.idPermiso = idPermiso;
    }

    public String getDescripcion() {
        return this.descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }

    public int getActivo() {
        return this.activo;
    }

    public void setActivo(int activo) {
        this.activo = activo;
    }

    @Override
    public String toString() {
        return "idPermiso :: " + idPermiso +
                " idPermiso:: " + descripcion +
                " activo:: " + activo;
    }
}
