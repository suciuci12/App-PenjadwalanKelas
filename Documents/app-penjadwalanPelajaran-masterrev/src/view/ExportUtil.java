/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package view;

/**
/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */

// ExportUtil.java
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.util.Date;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.SwingUtilities;
import javax.swing.table.TableModel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class ExportUtil {

    private ExportUtil() {}

     public static void exportTableToExcel(JTable table, Component parent,
                                          String nis, String nama, String kelas, String jurusan) {
        try {
            // ===== 1) Lokasi & nama file unik =====
            File docs = new File(System.getProperty("user.home"), "Documents");
            String ts = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
            File defaultFile = new File(docs, "jadwal_" + ts + ".xlsx");

            // ===== 2) Dialog simpan =====
            JFileChooser chooser = new JFileChooser(docs);
            chooser.setDialogTitle("Simpan sebagai Excel");
            chooser.setSelectedFile(defaultFile);
            if (chooser.showSaveDialog(SwingUtilities.getWindowAncestor(parent)) != JFileChooser.APPROVE_OPTION) return;

            File out = chooser.getSelectedFile();
            if (!out.getName().toLowerCase().endsWith(".xlsx")) {
                out = new File(out.getParentFile(), out.getName() + ".xlsx");
            }

            // ===== 3) Tulis workbook =====
            try (Workbook wb = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(out)) {
                Sheet sheet = wb.createSheet("Jadwal");
                TableModel m = table.getModel();
                int colCount = Math.max(1, m.getColumnCount());
                int rowCount = m.getRowCount();

                // ---------- Fonts ----------
                Font boldFont  = wb.createFont();  boldFont.setBold(true);
                Font titleFont = wb.createFont();  titleFont.setBold(true); titleFont.setFontHeightInPoints((short)14);

                // helper border
                java.util.function.Consumer<CellStyle> withBorder = st -> {
                    st.setBorderTop(BorderStyle.THIN);
                    st.setBorderBottom(BorderStyle.THIN);
                    st.setBorderLeft(BorderStyle.THIN);
                    st.setBorderRight(BorderStyle.THIN);
                };

                // ---------- Styles ----------
                // Info siswa (tanpa border)
                CellStyle labelStyle = wb.createCellStyle();
                labelStyle.setFont(boldFont);

                CellStyle colonStyle = wb.createCellStyle();
                colonStyle.setAlignment(HorizontalAlignment.CENTER);

                CellStyle valueStyle = wb.createCellStyle();
                valueStyle.setAlignment(HorizontalAlignment.LEFT);

                // Title (tanpa border)
                CellStyle titleStyle = wb.createCellStyle();
                titleStyle.setFont(titleFont);
                titleStyle.setAlignment(HorizontalAlignment.CENTER);

                // Header tabel (dengan border)
                CellStyle headerStyle = wb.createCellStyle();
                headerStyle.setFont(boldFont);
                headerStyle.setAlignment(HorizontalAlignment.CENTER);
                headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                withBorder.accept(headerStyle);

                // Data tabel – kiri (dengan border)
                CellStyle textStyle = wb.createCellStyle();
                textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                textStyle.setAlignment(HorizontalAlignment.LEFT);
                withBorder.accept(textStyle);

                // Data tabel – tengah (dengan border)
                CellStyle centerStyle = wb.createCellStyle();
                centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                centerStyle.setAlignment(HorizontalAlignment.CENTER);
                withBorder.accept(centerStyle);

                // Data tabel – waktu (dengan border + format jam)
                DataFormat df = wb.createDataFormat();
                CellStyle timeStyle = wb.createCellStyle();
                timeStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                timeStyle.setAlignment(HorizontalAlignment.CENTER);
                timeStyle.setDataFormat(df.getFormat("HH:mm:ss"));
                withBorder.accept(timeStyle);

                // ========== INFO SISWA ==========
                String[][] infos = {
                    {"NIS", nis},
                    {"Nama", nama},
                    {"Kelas", kelas},
                    {"Jurusan", jurusan}
                };

                int r = 0;
                for (String[] it : infos) {
                    Row row = sheet.createRow(r++);
                    // kolom A = label
                    Cell c0 = row.createCell(0);
                    c0.setCellValue(it[0]);
                    c0.setCellStyle(labelStyle);

                    // kolom B = ":"
                    Cell c1 = row.createCell(1);
                    c1.setCellValue(":");
                    c1.setCellStyle(colonStyle);

                    // kolom C = value (merge sampai kol terakhir)
                    Cell c2 = row.createCell(2);
                    c2.setCellValue(it[1]);
                    c2.setCellStyle(valueStyle);

                    int lastCol = Math.max(2, colCount - 1);
                    sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 2, lastCol));
                }
                // kecilkan kolom “:” agar rapat
                sheet.setColumnWidth(1, 3 * 256);

                r++; // baris kosong

                // ========== TITLE ==========
                Row titleRow = sheet.createRow(r++);
                Cell titleCell = titleRow.createCell(0);
                titleCell.setCellValue("JADWAL MATA PELAJARAN");
                titleCell.setCellStyle(titleStyle);
                sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, colCount - 1));
                titleRow.setHeightInPoints(20);

                // ========== HEADER ==========
                Row header = sheet.createRow(r++);
                header.setHeightInPoints(18);
                for (int c = 0; c < colCount; c++) {
                    Cell cell = header.createCell(c);
                    cell.setCellValue(m.getColumnName(c));
                    cell.setCellStyle(headerStyle); // <-- border di header
                }

                // ========== DATA ==========
                for (int i = 0; i < rowCount; i++) {
                    Row row = sheet.createRow(r++);
                    row.setHeightInPoints(16);
                    for (int j = 0; j < colCount; j++) {
                        Object v = m.getValueAt(i, j);
                        Cell cell = row.createCell(j);

                        // anggap dua kolom terakhir adalah jam (Jam Mulai, Jam Selesai)
                        if ((j == colCount - 2 || j == colCount - 1) && v != null) {
                            try {
                                String s = v.toString();
                                if (s.length() == 5) s = s + ":00"; // HH:mm → HH:mm:ss
                                LocalTime t = LocalTime.parse(s);
                                cell.setCellValue(t.toSecondOfDay() / 86400.0);
                                cell.setCellStyle(timeStyle);        // <-- border + format jam
                                continue;
                            } catch (Exception ignore) { /* fallback ke text */ }
                        }

                        cell.setCellValue(v == null ? "" : v.toString());
                        // kolom pertama (mapel) rata kiri, sisanya tengah
                        cell.setCellStyle(j == 0 ? textStyle : centerStyle);  // <-- semua data ada border
                    }
                }

                // Auto size kolom
                for (int c = 0; c < colCount; c++) {
                    sheet.autoSizeColumn(c, true);
                    int w = sheet.getColumnWidth(c);
                    sheet.setColumnWidth(c, Math.min(255*256, w + 800)); // sedikit padding
                }

                wb.write(fos);
            }

            JOptionPane.showMessageDialog(parent, "Excel tersimpan:\n" + out.getAbsolutePath());

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent, "Gagal export: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            ex.printStackTrace();
        }
    }
}