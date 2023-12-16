package com.example.prueba_excel;

import java.util.List;

public record PerfilPermisosDTO(Integer idPerfil,
        String descripcion,
        Integer activo,
        List<PermisoDTO> permisos) {
}
