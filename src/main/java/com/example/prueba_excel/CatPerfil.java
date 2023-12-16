package com.example.prueba_excel;

import java.util.List;

public class CatPerfil {
    private int idPerfil;
    private String descripcion;
    private int activo;
    private List<CatPermiso> permisos;

    public CatPerfil() {
        super();
    }

    public CatPerfil(int idPerfil, String descripcion, int activo) {
        super();
        this.idPerfil = idPerfil;
        this.descripcion = descripcion;
        this.activo = activo;
    }

    public CatPerfil(PerfilPermisosDTO perfilPermisosDTO) {
        idPerfil = perfilPermisosDTO.idPerfil();
        descripcion = perfilPermisosDTO.descripcion();
        activo = perfilPermisosDTO.activo();
        permisos = perfilPermisosDTO.permisos().stream().map(CatPermiso::new).toList();
    }

    public List<CatPermiso> getPermisos() {
        return permisos;
    }

    public int getIdPerfil() {
        return idPerfil;
    }

    public void setIdPerfil(int idPerfil) {
        this.idPerfil = idPerfil;
    }

    public String getDescripcion() {
        return descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }

    public int getActivo() {
        return activo;
    }

    public void setActivo(int activo) {
        this.activo = activo;
    }

    @Override
    public String toString() {
        return "idPerfil :: " + idPerfil +
                " idPermiso:: " + descripcion +
                " activo:: " + activo;
    }
}
