package com.example.prueba_excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
public class PruebaExcelController {
    @GetMapping("/excel")
    public ResponseEntity<byte[]> generarReporteExcel() {

        try (HSSFWorkbook workbook = new HSSFWorkbook()) { // Utilizando HSSFWorkbook para el formato antiguo
            Sheet sheet = workbook.createSheet("Report");

            // Datos ficticios
            String[] headers = { "Nombre", "Apellido", "Edad" };
            Object[][] data = {
                    { "John", "Doe", 25 },
                    { "Jane", "Smith", 30 },
                    { "Bob", "Johnson", 22 }
            };

            // Crear fila de encabezados
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Llenar los datos
            for (int rowIndex = 0; rowIndex < data.length; rowIndex++) {
                Row row = sheet.createRow(rowIndex + 1);
                for (int cellIndex = 0; cellIndex < data[rowIndex].length; cellIndex++) {
                    Cell cell = row.createCell(cellIndex);
                    if (data[rowIndex][cellIndex] instanceof String) {
                        cell.setCellValue((String) data[rowIndex][cellIndex]);
                    } else if (data[rowIndex][cellIndex] instanceof Integer) {
                        cell.setCellValue((Integer) data[rowIndex][cellIndex]);
                    }
                }
            }

            // Convertir el libro de trabajo a un arreglo de bytes
            try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
                workbook.write(byteArrayOutputStream);

                HttpHeaders headersResponse = new HttpHeaders();
                headersResponse.set("Content-Type", "application/vnd.ms-excel;charset=UTF-8");
                headersResponse.setContentDispositionFormData("attachment", "report.xls"); // Nota: Cambiado a extensión
                                                                                           // .xls

                return ResponseEntity.ok().headers(headersResponse).body(byteArrayOutputStream.toByteArray());
            }
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body(null);
        }
    }

    @GetMapping("/exportarExcel")
    public ResponseEntity<byte[]> exportarExcel() {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            String[] perfiles = { "Administrador", "Usuario", "Invitado" };

            for (int i = 0; i < perfiles.length; i++) {
                CatPerfil perfil = new CatPerfil(i + 1, perfiles[i], 1);

                HSSFSheet sheet = workbook.createSheet(perfil.getDescripcion());
                HSSFCellStyle cellStyle = workbook.createCellStyle();
                setText(getCell(sheet, 1, 0), "Perfil: " + perfil.getDescripcion());
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                HSSFFont fuente = workbook.createFont();
                fuente.setBold(true);
                cellStyle.setFont(fuente);
                getCell(sheet, 1, 0).setCellStyle(cellStyle);
                sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 1));
                List<CatPermiso> listaPermiso = new ArrayList<>();
                setText(getCell(sheet, 4, 0), "Id");
                setText(getCell(sheet, 4, 1), "Descripcion");
                getCell(sheet, 4, 0).setCellStyle(cellStyle);
                getCell(sheet, 4, 1).setCellStyle(cellStyle);

                setText(getCell(sheet, 2, 0), "Permisos que pertenecen al perfil");
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                fuente = workbook.createFont();
                fuente.setColor(IndexedColors.DARK_BLUE.getIndex());
                cellStyle.setFont(fuente);
                getCell(sheet, 2, 0).setCellStyle(cellStyle);
                sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));

                HSSFCellStyle style = workbook.createCellStyle();
                style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setAlignment(HorizontalAlignment.LEFT);
                fuente = workbook.createFont();
                fuente.setColor(IndexedColors.DARK_BLUE.getIndex());
                style.setFont(fuente);
                /*******************************************/
                for (int j = 0; j < listaPermiso.size(); j++) {
                    CatPermiso permiso = listaPermiso.get(j);
                    getCell(sheet, j + 5, 0).setCellType(CellType.NUMERIC);
                    getCell(sheet, j + 5, 0).setCellValue(permiso.getIdPermiso());
                    setText(getCell(sheet, j + 5, 1), permiso.getDescripcion());
                    getCell(sheet, j + 5, 0).setCellStyle(cellStyle);
                    getCell(sheet, j + 5, 1).setCellStyle(style);
                }

                getCell(sheet, listaPermiso.size() + 5, 0).setCellValue("Total :" + listaPermiso.size());
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                getCell(sheet, listaPermiso.size() + 5, 0).setCellStyle(cellStyle);
                sheet.addMergedRegion(new CellRangeAddress(listaPermiso.size() + 5, listaPermiso.size() + 5, 0, 1));

                for (int k = 0; k <= 2; k++) {
                    sheet.autoSizeColumn((short) 1);
                }

            }

            try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
                workbook.write(byteArrayOutputStream);

                HttpHeaders headersResponse = new HttpHeaders();
                headersResponse.set("Content-Type", "application/vnd.ms-excel;charset=UTF-8");
                headersResponse.setContentDispositionFormData("attachment", "report.xls"); // Nota: Cambiado a extensión
                                                                                           // .xls

                return ResponseEntity.ok().headers(headersResponse).body(byteArrayOutputStream.toByteArray());
            }
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body(null);
        }
    }

    protected void setText(HSSFCell cell, String text) {
        cell.setCellType(CellType.STRING);
        cell.setCellValue(text);
    }

    protected HSSFCell getCell(HSSFSheet sheet, int row, int col) {
        HSSFRow sheetRow = sheet.getRow(row);
        if (sheetRow == null) {
            sheetRow = sheet.createRow(row);
        }
        HSSFCell cell = sheetRow.getCell((short) col);
        if (cell == null) {
            cell = sheetRow.createCell((short) col);
        }
        return cell;
    }
}
