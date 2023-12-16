package com.example.prueba_excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class PruebaExcel2Controller {
    @GetMapping("/exportarExcel2")
    public ResponseEntity<byte[]> exportarExcel() throws IOException {
        HttpHeaders headersResponse = new HttpHeaders();
        headersResponse.set("Content-Type", "application/vnd.ms-excel;charset=UTF-8");
        headersResponse.setContentDispositionFormData("attachment", "report.xls"); // Nota: Cambiado a extensión
        return new ResponseEntity<>(generarReporteExcelPerfilPermisos(),
                headersResponse, HttpStatus.OK);
    }

    public byte[] generarReporteExcelPerfilPermisos() throws IOException {
        HSSFWorkbook libroTrabajo = new HSSFWorkbook();

        List<PerfilPermisosDTO> perfilPermisosDTOs = new ArrayList<>();

        IntStream.range(0, 5).forEach(i -> {
            List<PermisoDTO> permisoDTOs = new ArrayList<>();

            IntStream.range(0, 99).forEach(j -> {
                PermisoDTO permisoDTO = new PermisoDTO(j, "Permiso " + j, 1);
                permisoDTOs.add(permisoDTO);
            });

            PerfilPermisosDTO perfilPermisosDTO = new PerfilPermisosDTO(i, "Perfil " + i, 1, permisoDTOs);
            perfilPermisosDTOs.add(perfilPermisosDTO);
        });

        HSSFCellStyle estiloEncabezado = crearEstiloEncabezado(libroTrabajo);
        HSSFCellStyle estiloCeldaBlanca = crearEstiloCeldaBlanca(libroTrabajo);
        HSSFCellStyle estiloCeldaAzul = crearEstiloCeldaAzul(libroTrabajo);
        HSSFCellStyle estiloCeldaBlancaColPermisoId = crearEstiloCeldaBlanca(libroTrabajo);
        HSSFCellStyle estiloCeldaBlancaColPermisoDescripcion = crearEstiloCeldaBlanca(libroTrabajo);

        perfilPermisosDTOs.stream().forEach(perfil -> generarHojaExcel(
                new CatPerfil(perfil), libroTrabajo, estiloEncabezado, estiloCeldaBlanca, estiloCeldaAzul,
                estiloCeldaBlancaColPermisoId, estiloCeldaBlancaColPermisoDescripcion));

        ByteArrayOutputStream flujoSalidaBytes = new ByteArrayOutputStream();
        libroTrabajo.write(flujoSalidaBytes);

        return flujoSalidaBytes.toByteArray();

    }

    /**
     * Genera la hoja de Excel que contiene los permisos del perfil correspondiente.
     *
     * @param perfil       El perfil del cual se desean obtener los permisos.
     * @param libroTrabajo El libro de trabajo que contiene la hoja de Excel.
     */
    public void generarHojaExcel(CatPerfil perfil, HSSFWorkbook libroTrabajo, HSSFCellStyle estiloEncabezado,
            HSSFCellStyle estiloCeldaBlanca, HSSFCellStyle estiloCeldaAzul, HSSFCellStyle estiloCeldaBlancaColPermisoId,
            HSSFCellStyle estiloCeldaBlancaColPermisoDescripcion) {
        List<CatPermiso> permisos = perfil.getPermisos();
        HSSFSheet hoja = libroTrabajo.createSheet(perfil.getDescripcion());

        estiloCeldaBlanca.setAlignment(HorizontalAlignment.CENTER);

        establecerTextoCelda(obtenerCelda(hoja, 1, 0), "Perfil: " + perfil.getDescripcion());
        obtenerCelda(hoja, 1, 0).setCellStyle(estiloEncabezado);
        hoja.addMergedRegion(new CellRangeAddress(1, 1, 0, 1));

        establecerTextoCelda(obtenerCelda(hoja, 4, 0), "Id");
        establecerTextoCelda(obtenerCelda(hoja, 4, 1), "Descripcion");
        obtenerCelda(hoja, 4, 0).setCellStyle(estiloEncabezado);
        obtenerCelda(hoja, 4, 1).setCellStyle(estiloEncabezado);

        establecerTextoCelda(obtenerCelda(hoja, 2, 0), "Permisos que pertenecen al perfil");
        obtenerCelda(hoja, 2, 0).setCellStyle(estiloCeldaBlanca);
        hoja.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));

        IntStream.range(0, permisos.size()).forEach(i -> {
            CatPermiso permiso = permisos.get(i);

            obtenerCelda(hoja, i + 5, 0).setCellType(CellType.NUMERIC);
            obtenerCelda(hoja, i + 5, 0).setCellValue(permiso.getIdPermiso());
            estiloCeldaBlancaColPermisoId.setAlignment(HorizontalAlignment.CENTER);
            obtenerCelda(hoja, i + 5, 0).setCellStyle(estiloCeldaBlancaColPermisoId);

            establecerTextoCelda(obtenerCelda(hoja, i + 5, 1), permiso.getDescripcion());
            estiloCeldaBlancaColPermisoDescripcion.setAlignment(HorizontalAlignment.LEFT);
            obtenerCelda(hoja, i + 5, 1).setCellStyle(estiloCeldaBlancaColPermisoDescripcion);
        });

        obtenerCelda(hoja, permisos.size() + 5, 0).setCellValue("Total: " + permisos.size());
        obtenerCelda(hoja, permisos.size() + 5, 0).setCellStyle(estiloCeldaAzul);
        hoja.addMergedRegion(new CellRangeAddress(permisos.size() + 5, permisos.size() + 5, 0, 1));

        IntStream.range(0, 2).forEach(columna -> hoja.autoSizeColumn(1));
    }

    /**
     * Crea el estilo de la celda que contiene el encabezado de la hoja de Excel.
     *
     * @param libroTrabajo El libro de trabajo que contiene la hoja de Excel.
     * @return El estilo de la celda que contiene el encabezado de la hoja de Excel.
     */
    public HSSFCellStyle crearEstiloEncabezado(HSSFWorkbook libroTrabajo) {
        HSSFCellStyle estiloCelda = libroTrabajo.createCellStyle();
        estiloCelda.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estiloCelda.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont fuenteNegrita = libroTrabajo.createFont();
        fuenteNegrita.setBold(true);
        estiloCelda.setFont(fuenteNegrita);

        return estiloCelda;
    }

    /**
     * Crea el estilo de una celda con fondo blanco.
     *
     * @param libroTrabajo El libro de trabajo que contiene la hoja de Excel.
     * @return El estilo de la celda con fondo blanco.
     */
    public HSSFCellStyle crearEstiloCeldaBlanca(HSSFWorkbook libroTrabajo) {
        HSSFCellStyle estiloCelda = libroTrabajo.createCellStyle();
        estiloCelda.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        HSSFFont fuente = libroTrabajo.createFont();
        fuente.setColor(IndexedColors.DARK_BLUE.getIndex());
        estiloCelda.setFont(fuente);

        return estiloCelda;
    }

    /**
     * Crea el estilo de una celda con fondo azul.
     *
     * @param libroTrabajo El libro de trabajo que contiene la hoja de Excel.
     * @return El estilo de la celda con fondo azul.
     */
    public HSSFCellStyle crearEstiloCeldaAzul(HSSFWorkbook libroTrabajo) {
        HSSFCellStyle estiloCelda = libroTrabajo.createCellStyle();
        estiloCelda.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        estiloCelda.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estiloCelda.setAlignment(HorizontalAlignment.CENTER);

        return estiloCelda;
    }

    /**
     * Establece el texto de la celda correspondiente.
     *
     * @param celda Celda a la cual se le establece el texto.
     * @param texto Texto que se establece en la celda.
     */
    public void establecerTextoCelda(HSSFCell celda, String texto) {
        celda.setCellType(CellType.STRING);
        celda.setCellValue(texto);
    }

    /**
     * Obtiene la celda correspondiente a la fila y columna especificadas.
     *
     * @param hoja    La hoja de Excel que contiene la celda.
     * @param fila    La fila de la celda.
     * @param columna La columna de la celda.
     * @return La celda correspondiente a la fila y columna especificadas.
     */
    public HSSFCell obtenerCelda(HSSFSheet hoja, int fila, int columna) {
        HSSFRow filaHoja = hoja.getRow(fila);

        if (filaHoja == null) {
            filaHoja = hoja.createRow(fila);
        }

        HSSFCell celda = filaHoja.getCell(columna);

        if (celda == null) {
            celda = filaHoja.createCell(columna);
        }

        return celda;
    }
}
