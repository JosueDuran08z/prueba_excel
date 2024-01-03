package com.example.prueba_excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;
import java.awt.GraphicsEnvironment;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class PruebaExcel2Controller {
    @GetMapping("/excel2")
    public ResponseEntity<byte[]> exportarExcel() throws IOException {
        HttpHeaders headersResponse = new HttpHeaders();
        headersResponse.set("Content-Type", "application/vnd.ms-excel;charset=UTF-8");
        headersResponse.setContentDispositionFormData("attachment", "report.xls"); // Nota: Cambiado a extensión
        return new ResponseEntity<>(generarReporteExcelPerfilPermisos(),
                headersResponse, HttpStatus.OK);
    }

    public byte[] generarReporteExcelPerfilPermisos() throws IOException {
        HSSFWorkbook libroTrabajo = new HSSFWorkbook();
        Sheet sheet = libroTrabajo.createSheet("Hoja1");
        // Crear una fila en la hoja
        Row row = sheet.createRow(0);

        // Crear celdas en A1 y B1
        Cell cellA1 = row.createCell(0);
        Cell cellB1 = row.createCell(1);

        // Crear un objeto para unir celdas (A1:B1)
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 1);
        sheet.addMergedRegion(region);

        // Crear un estilo y establecer el color de fondo blanco
        CellStyle style = libroTrabajo.createCellStyle();
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Aplicar el estilo a las celdas fusionadas
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            Row rowRegion = CellUtil.getRow(i, sheet);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                Cell cellRegion = CellUtil.getCell(rowRegion, j);
                cellRegion.setCellStyle(style);
            }
        }

        // Crear una nueva fila para A2 y B2
        Row row2 = sheet.createRow(1);

        // Crear celdas en A2 y B2
        Cell cellA2 = row2.createCell(0);
        Cell cellB2 = row2.createCell(1);

        // Establecer valores y estilo en A2 y B2
        cellA2.setCellValue("Perfil: DESARROLLO");
        cellB2.setCellValue(""); // No contenido en B2
        CellRangeAddress regionA2B2 = new CellRangeAddress(1, 1, 0, 1);
        sheet.addMergedRegion(regionA2B2);

        CellStyle styleA2B2 = libroTrabajo.createCellStyle();
        styleA2B2.setAlignment(HorizontalAlignment.CENTER);
        styleA2B2.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        styleA2B2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font fontA2B2 = libroTrabajo.createFont();
        fontA2B2.setBold(true);
        styleA2B2.setFont(fontA2B2);

        Row row3 = sheet.createRow(2);
        Cell cellA3 = row3.createCell(0);

        // Establecer valores y configuraciones en las celdas A3 y B3
        cellA3.setCellValue("Permisos que pertenecen al perfil");

        // Crear un objeto para unir celdas (A3:B3)
        CellRangeAddress regionA3B3 = new CellRangeAddress(2, 2, 0, 1);
        sheet.addMergedRegion(regionA3B3);

        // Crear un estilo para las celdas A3 y B3
        CellStyle styleA3B3 = libroTrabajo.createCellStyle();
        styleA3B3.setAlignment(HorizontalAlignment.CENTER);
        styleA3B3.setVerticalAlignment(VerticalAlignment.CENTER);
        styleA3B3.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        styleA3B3.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Configurar el color del texto en las celdas A3 y B3
        Font fontA3B3 = libroTrabajo.createFont();
        fontA3B3.setColor(IndexedColors.DARK_BLUE.getIndex());
        styleA3B3.setFont(fontA3B3);

        // Aplicar el estilo a las celdas fusionadas (A3:B3)
        for (int i = regionA3B3.getFirstRow(); i <= regionA3B3.getLastRow(); i++) {
            Row rowRegion = CellUtil.getRow(i, sheet);
            for (int j = regionA3B3.getFirstColumn(); j <= regionA3B3.getLastColumn(); j++) {
                Cell cellRegion = CellUtil.getCell(rowRegion, j);
                cellRegion.setCellStyle(styleA3B3);
            }
        }

        Row row5 = sheet.createRow(4);
        Cell cellA5 = row5.createCell(0);
        Cell cellB5 = row5.createCell(1);

        // Establecer valores y configuraciones en las celdas A5 y B5
        cellA5.setCellValue("Id");
        cellB5.setCellValue("Descripcion");

        // Aplicar el estilo de A2 a A5 y B5
        cellA5.setCellStyle(styleA2B2);
        cellB5.setCellStyle(styleA2B2);

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

        // Crear un estilo específico para las celdas de la columna B6
        CellStyle styleColumnaB = libroTrabajo.createCellStyle();
        styleColumnaB.setAlignment(HorizontalAlignment.LEFT);
        styleColumnaB.setVerticalAlignment(VerticalAlignment.CENTER);
        styleColumnaB.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        styleColumnaB.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Configurar el color del texto en las celdas de la columna A6
        Font fontColumnaA = libroTrabajo.createFont();
        fontColumnaA.setColor(IndexedColors.DARK_BLUE.getIndex());
        styleColumnaB.setFont(fontColumnaA);

        CellStyle estiloColumnaPerfil = libroTrabajo.createCellStyle();
        estiloColumnaPerfil.setAlignment(HorizontalAlignment.LEFT);
        estiloColumnaPerfil.setVerticalAlignment(VerticalAlignment.CENTER);
        estiloColumnaPerfil.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        estiloColumnaPerfil.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estiloColumnaPerfil.setFont(fontColumnaA);

        // Iterar sobre la lista de permisos e insertar datos en el archivo Excel
        int rowNum = 5; // Empezar desde la fila 6

        PerfilPermisosDTO perfil = perfilPermisosDTOs.get(0);
        List<PermisoDTO> permisos = perfil.permisos();

        for (PermisoDTO permiso : permisos) {
            Row renglonPermiso = sheet.createRow(rowNum);

            // Celda A6 (Id)
            Cell cellA6 = renglonPermiso.createCell(0);
            cellA6.setCellValue(permiso.idPermiso());
            cellA6.setCellStyle(styleColumnaB);

            // Celda B6 (Descripcion)

            Cell cellB6 = renglonPermiso.createCell(1);
            cellB6.setCellValue("AC_CONSULTA_SERIES_CARGADAS_MODF GOKU 123ASDAS ASDAS ASDSAD SDS__2323");
            cellB6.setCellStyle(estiloColumnaPerfil);

            rowNum++;
        }

        CellStyle estiloCeldaTotal = libroTrabajo.createCellStyle();
        estiloCeldaTotal.setAlignment(HorizontalAlignment.CENTER);
        estiloCeldaTotal.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        estiloCeldaTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Calcular y mostrar el total de permisos al final
        int totalPermisos = permisos.size(); // Moverse a la siguiente fila después de la lista
        Row totalRow = sheet.createRow(rowNum);

        // Fusionar celdas A(totalRow):B(totalRow)
        CellRangeAddress totalRegion = new CellRangeAddress(rowNum, rowNum, 0, 1);
        sheet.addMergedRegion(totalRegion);

        // Crear celda para mostrar el total
        Cell totalCell = totalRow.createCell(0);
        totalCell.setCellValue("Total de Permisos: " + totalPermisos);
        totalCell.setCellStyle(estiloCeldaTotal);

        // Crear celda para mostrar el valor del total
        Cell totalValueCell = totalRow.createCell(1);
        totalValueCell.setCellValue(totalPermisos);
        totalValueCell.setCellStyle(estiloCeldaTotal);
        applyStyleToRegion(sheet, regionA2B2, styleA2B2);
        int longitud = "".length() + 8;
        // sheet.setColumnWidth(1, longitud * 256);
        for (int i = 0; i <= 2; i++) {
            sheet.autoSizeColumn(1, true); // El segundo parámetro indica ajustar solo el ancho visible
        }
        ByteArrayOutputStream flujoSalidaBytes = new ByteArrayOutputStream();
        libroTrabajo.write(flujoSalidaBytes);

        return flujoSalidaBytes.toByteArray();

    }

    private static void applyStyleToRegion(Sheet sheet, CellRangeAddress region, CellStyle style) {
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            Row rowRegion = CellUtil.getRow(i, sheet);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                Cell cellRegion = CellUtil.getCell(rowRegion, j);
                cellRegion.setCellStyle(style);
            }
        }
    }
}
