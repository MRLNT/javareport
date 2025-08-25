import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

// Hapus deklarasi 'package'
public class UploadProcessor {

    // --- Helper untuk formatting cell ---
    private static String cellStr(DataFormatter fmt, Cell c) {
        return (c == null) ? "" : fmt.formatCellValue(c).trim();
    }
    private static boolean notEmpty(String s) { return s != null && !s.trim().isEmpty(); }

    // --- Struktur agregat per PASS ---
    private static class PassAgg {
        Map<String, Boolean> bca = new LinkedHashMap<>();
        Map<String, Boolean> vendor = new LinkedHashMap<>();
        boolean statusYes = false;
        String visitDate = "";
        String activity  = "";
        String purpose   = "";
        String pic       = "";
        String assetInc  = "";
    }

    public void processExcel(String inputFilePath, String templateFilePath, String outputFilePath) throws IOException {

        try (InputStream templateStream = new FileInputStream(templateFilePath);
             Workbook templateWorkbook = new XSSFWorkbook(templateStream)) {

            Sheet targetSheet = templateWorkbook.getSheetAt(0);

            try (InputStream sourceInputStream = new FileInputStream(inputFilePath);
                 Workbook sourceWorkbook = new XSSFWorkbook(sourceInputStream)) {

                Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
                DataFormatter fmt = new DataFormatter();

                Map<String, PassAgg> byPass = new LinkedHashMap<>();

                for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
                    Row r = sourceSheet.getRow(i);
                    if (r == null) continue;

                    String passId = cellStr(fmt, r.getCell(0));
                    if (!notEmpty(passId)) continue;

                    PassAgg agg = byPass.computeIfAbsent(passId, k -> new PassAgg());

                    String st = cellStr(fmt, r.getCell(1)).toUpperCase(Locale.ROOT);
                    if ("OPENED".equals(st) || "CLOSED".equals(st)) agg.statusYes = true;

                    String activity = cellStr(fmt, r.getCell(2));
                    String assetInc = cellStr(fmt, r.getCell(3));
                    String pic      = cellStr(fmt, r.getCell(4));
                    String purpose  = cellStr(fmt, r.getCell(5));
                    String vdate    = cellStr(fmt, r.getCell(6));

                    if (!notEmpty(agg.activity) && notEmpty(activity)) agg.activity = activity;
                    if (!notEmpty(agg.assetInc) && notEmpty(assetInc)) agg.assetInc = assetInc;
                    if (!notEmpty(agg.pic)      && notEmpty(pic))      agg.pic      = pic;
                    if (!notEmpty(agg.purpose)  && notEmpty(purpose))  agg.purpose  = purpose;
                    if (!notEmpty(agg.visitDate)&& notEmpty(vdate))    agg.visitDate= vdate;

                    String company = cellStr(fmt, r.getCell(7)).toLowerCase(Locale.ROOT);
                    String name    = cellStr(fmt, r.getCell(8));
                    String checkIn = cellStr(fmt, r.getCell(9));
                    boolean hasCheckIn = notEmpty(checkIn) && !checkIn.equals("-");

                    if (notEmpty(name)) {
                        if (company.contains("bank central asia")) {
                            agg.bca.put(name, Boolean.valueOf(hasCheckIn || agg.bca.getOrDefault(name, false)));
                        } else {
                            agg.vendor.put(name, Boolean.valueOf(hasCheckIn || agg.vendor.getOrDefault(name, false)));
                        }
                    }
                }

                int rowIdx = 10;
                int totalPass = 0;
                int totalRealisasi = 0;
                int totalBcaCeklis = 0;
                int totalVendorCeklis = 0;
                int totalTidak = 0;
                int totalBcaNoCeklis = 0;
                int totalVendorNoCeklis = 0;

                for (Map.Entry<String, PassAgg> e : byPass.entrySet()) {
                    String passId = e.getKey();
                    PassAgg agg = e.getValue();

                    Row row = targetSheet.getRow(rowIdx);
                    if (row == null) row = targetSheet.createRow(rowIdx);

                    StringBuilder bcaNames = new StringBuilder();
                    for (Map.Entry<String, Boolean> n : agg.bca.entrySet()) {
                        bcaNames.append(n.getKey());
                        if (Boolean.TRUE.equals(n.getValue())) {
                            bcaNames.append("✅");
                            totalBcaCeklis++;
                        } else {
                            totalBcaNoCeklis++;
                        }
                        bcaNames.append("\n");
                    }
                    StringBuilder vendorNames = new StringBuilder();
                    for (Map.Entry<String, Boolean> n : agg.vendor.entrySet()) {
                        vendorNames.append(n.getKey());
                        if (Boolean.TRUE.equals(n.getValue())) {
                            vendorNames.append("✅");
                            totalVendorCeklis++;
                        } else {
                            totalVendorNoCeklis++;
                        }
                        vendorNames.append("\n");
                    }

                    Cell cA = row.getCell(0); if (cA == null) cA = row.createCell(0); cA.setCellValue(bcaNames.toString().trim());
                    Cell cB = row.getCell(1); if (cB == null) cB = row.createCell(1); cB.setCellValue(agg.bca.size());
                    Cell cC = row.getCell(2); if (cC == null) cC = row.createCell(2); cC.setCellValue(vendorNames.toString().trim());
                    Cell cD = row.getCell(3); if (cD == null) cD = row.createCell(3); cD.setCellValue(agg.vendor.size());
                    Cell cE = row.getCell(4); if (cE == null) cE = row.createCell(4); cE.setCellValue(agg.statusYes ? "Ya" : "Tidak");
                    Cell cF = row.getCell(5); if (cF == null) cF = row.createCell(5); cF.setCellValue(agg.visitDate);
                    Cell cH = row.getCell(7); if (cH == null) cH = row.createCell(7); cH.setCellValue(passId);
                    Cell cI = row.getCell(8); if (cI == null) cI = row.createCell(8); cI.setCellValue(agg.pic);
                    Cell cJ = row.getCell(9); if (cJ == null) cJ = row.createCell(9); cJ.setCellValue(agg.activity);
                    Cell cK = row.getCell(10); if (cK == null) cK = row.createCell(10); cK.setCellValue(agg.purpose);
                    Cell cS = row.getCell(18); if (cS == null) cS = row.createCell(18); cS.setCellValue(agg.assetInc);

                    totalPass++;
                    if (agg.statusYes) totalRealisasi++;
                    else totalTidak++;

                    rowIdx++;
                }

                Row row2 = targetSheet.getRow(1); if (row2 == null) row2 = targetSheet.createRow(1); Cell d2 = row2.getCell(3); if (d2 == null) d2 = row2.createCell(3); d2.setCellValue(totalPass);
                Row row3 = targetSheet.getRow(2); if (row3 == null) row3 = targetSheet.createRow(2); Cell d3 = row3.getCell(3); if (d3 == null) d3 = row3.createCell(3); d3.setCellValue(totalRealisasi);
                Row row4 = targetSheet.getRow(3); if (row4 == null) row4 = targetSheet.createRow(3); Cell d4 = row4.getCell(3); if (d4 == null) d4 = row4.createCell(3); d4.setCellValue(totalBcaCeklis);
                Row row5 = targetSheet.getRow(4); if (row5 == null) row5 = targetSheet.createRow(4); Cell d5 = row5.getCell(3); if (d5 == null) d5 = row5.createCell(3); d5.setCellValue(totalVendorCeklis);
                Row row6 = targetSheet.getRow(5); if (row6 == null) row6 = targetSheet.createRow(5); Cell d6 = row6.getCell(3); if (d6 == null) d6 = row6.createCell(3); d6.setCellValue(totalTidak);
                Row row7 = targetSheet.getRow(6); if (row7 == null) row7 = targetSheet.createRow(6); Cell d7 = row7.getCell(3); if (d7 == null) d7 = row7.createCell(3); d7.setCellValue(totalBcaNoCeklis);
                Row row8 = targetSheet.getRow(7); if (row8 == null) row8 = targetSheet.createRow(7); Cell d8 = row8.getCell(3); if (d8 == null) d8 = row8.createCell(3); d8.setCellValue(totalVendorNoCeklis);

                sourceWorkbook.close();

                try (OutputStream out = new FileOutputStream(outputFilePath)) {
                    templateWorkbook.write(out);
                }
                System.out.println("Processing complete. Output saved to: " + outputFilePath);

            } catch (IOException e) {
                System.err.println("Error reading source Excel file: " + e.getMessage());
                throw e;
            }
        } catch (IOException e) {
            System.err.println("Error loading template Excel file: " + e.getMessage());
            throw e;
        }
    }
}