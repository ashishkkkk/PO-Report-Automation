package com.example.automation;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.yaml.snakeyaml.Yaml;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.ss.util.CellRangeAddress;
import java.nio.file.Paths;
import java.text.NumberFormat;
import java.util.ArrayList;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

@Command(name = "run-automation", mixinStandardHelpOptions = true, version = "0.1")
public class Main implements Runnable {

    @Option(names = {"-c", "--config"}, description = "Path to YAML config", required = false)
    private Path config;

    private static class AuartInfo {
        final String subHeading;
        final String item;

        AuartInfo(String subHeading, String item) {
            this.subHeading = subHeading;
            this.item = item;
        }
    }

    private static final Map<String, AuartInfo> AUART_MAPPING = new LinkedHashMap<>();
    static {
        // Vehicle
        AUART_MAPPING.put("ZMTB", new AuartInfo("MTO CS", "Vehicle"));
        AUART_MAPPING.put("ZMTO", new AuartInfo("MTO Dealer", "Vehicle"));
        AUART_MAPPING.put("ZMTS", new AuartInfo("MTS Dealer & CS", "Vehicle"));
        AUART_MAPPING.put("ZQTN", new AuartInfo("SVPO Dealer", "Vehicle"));
        AUART_MAPPING.put("ZVBL", new AuartInfo("Vector PO Dealer", "Vehicle"));
        AUART_MAPPING.put("ZUBB", new AuartInfo("Vector PO CS", "Vehicle"));
        // Spares
        AUART_MAPPING.put("FD", new AuartInfo("P2P PO Dealer", "SPARES"));
        AUART_MAPPING.put("ZNPD", new AuartInfo("NPO PO Dealer", "SPARES"));
        AUART_MAPPING.put("ZRSO", new AuartInfo("VOR PO Dealer", "SPARES"));
        AUART_MAPPING.put("ZSPA", new AuartInfo("Spares PO Dealer & CS", "SPARES"));
        AUART_MAPPING.put("ZBLK", new AuartInfo("Vector Spares PO Dealer", "SPARES"));
        // Oil
        AUART_MAPPING.put("ZOIL", new AuartInfo("Oil PO Dealer", "OIL"));
        AUART_MAPPING.put("ZOBL", new AuartInfo("Vector Oil PO Dealer", "OIL"));
        // Gear
        AUART_MAPPING.put("ZACW", new AuartInfo("Gear PO Dealer", "GEAR"));
        // GMA
        AUART_MAPPING.put("ZMCA", new AuartInfo("GMA PO Dealer", "GMA"));
        AUART_MAPPING.put("ZMTG", new AuartInfo("GMA PO Config Dealer", "GMA"));
        AUART_MAPPING.put("ZMYG", new AuartInfo("GMA SNOP Order Dealer", "GMA"));
        AUART_MAPPING.put("ZMCV", new AuartInfo("Vector GMA PO Dealer", "GMA"));
        AUART_MAPPING.put("ZUBZ", new AuartInfo("Vector GMA PO CS", "GMA"));
    }


    // Column Indexes (0-based)
    private static final int COL_DMSPONO = 6; // Column G
    private static final int COL_SALORD_NO = 13; // Column N
    private static final int COL_SPART = 4; // Column E - Assumed 'Spart' column, PLEASE VERIFY.

    public static void main(String[] args) {
        int exitCode = new CommandLine(new Main()).execute(args);
        System.exit(exitCode);
    }

    @Override
    public void run() {
        try {
            String inputPathStr = "D:\\Certificate\\DynamicsExport_639072487843937246.xlsx";
            String outputPathStr = "output.xlsx";
            String sapConnectionName = "PRE Load Balance";

            if (config != null) {
                if (!Files.exists(config)) throw new IllegalArgumentException("config file not found: " + config);
                Yaml yaml = new Yaml();
                try (InputStream in = Files.newInputStream(config)) {
                    @SuppressWarnings("unchecked")
                    Map<String, Object> cfg = (Map<String, Object>) yaml.load(in);
                    inputPathStr = cfg != null && cfg.get("input") != null ? cfg.get("input").toString() : null;
                    outputPathStr = cfg != null && cfg.get("output") != null ? cfg.get("output").toString() : "output.xlsx";
                    if (cfg != null && cfg.get("sap_connection") != null) sapConnectionName = cfg.get("sap_connection").toString();
                }
            }

            if (inputPathStr == null) {
                System.out.println("No input file specified. Please select the raw data file.");
                inputPathStr = promptForFile();
            }

            if (inputPathStr == null || inputPathStr.isEmpty()) {
                System.err.println("No input file selected. Exiting.");
                System.exit(1);
            }

            process(Paths.get(inputPathStr), Paths.get(outputPathStr), sapConnectionName);
            System.out.println("Automation completed successfully");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }

    private String promptForFile() {
        if (!java.awt.GraphicsEnvironment.isHeadless()) {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Select Raw Data Excel File");
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            fileChooser.setCurrentDirectory(new java.io.File("."));
            int result = fileChooser.showOpenDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                return fileChooser.getSelectedFile().getAbsolutePath();
            }
        }

        System.out.println("Enter the full path to the raw data Excel file:");
        Scanner scanner = new Scanner(System.in);
        if (scanner.hasNextLine()) {
            String path = scanner.nextLine().trim();
            return path.replace("\"", "");
        }
        return null;
    }

    private AuartInfo getAuartInfo(String auart, String spartName) {
        if (auart == null || auart.trim().isEmpty()) {
            return null;
        }
        auart = auart.trim();

        if ("ZUB".equals(auart)) {
            if ("Vehicle".equalsIgnoreCase(spartName)) return new AuartInfo("SVPO CS", "Vehicle");
            if ("Spares".equalsIgnoreCase(spartName)) return new AuartInfo("Spares PO Dealer & CS", "SPARES");
            if ("Oil".equalsIgnoreCase(spartName)) return new AuartInfo("Oil PO CS", "OIL");
            if ("Gear".equalsIgnoreCase(spartName)) return new AuartInfo("Gear PO CS", "GEAR");
            if ("GMA".equalsIgnoreCase(spartName)) return new AuartInfo("GMA PO CS", "GMA");
        }

        return AUART_MAPPING.get(auart);
    }

    public void process(Path inputPath, Path outputPath, String sapConnectionName) throws Exception {
        System.out.println("Starting PO Report Processing...");
        System.out.println("Reading from: " + inputPath);

        Map<String, String[]> allUniqueRecordsMap = new LinkedHashMap<>();
        List<String> headers = new ArrayList<>();
        int totalRowsRead = 0;

        // 1. Read and Split Data
        try (InputStream is = new FileInputStream(inputPath.toFile());
             Workbook workbook = StreamingReader.builder()
                     .rowCacheSize(100)
                     .bufferSize(4096)
                     .open(is)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                int lastCellNum = row.getLastCellNum();
                String[] rowData = new String[Math.max(lastCellNum, COL_SALORD_NO + 2)];

                for (int i = 0; i < lastCellNum; i++) {
                    Cell cell = row.getCell(i);
                    rowData[i] = (cell != null) ? cell.getStringCellValue() : "";
                }

                if (row.getRowNum() == 0) {
                    for (String h : rowData) {
                        if (h != null) headers.add(h);
                    }
                    continue;
                }

                totalRowsRead++;
                String sdn = (rowData.length > COL_SALORD_NO) ? rowData[COL_SALORD_NO] : "";
                String poNumber = (rowData.length > COL_DMSPONO) ? rowData[COL_DMSPONO] : "";

                if (poNumber != null && !poNumber.isEmpty()) {
                    // Smart Deduplication: Prefer records that have an SDN
                    if (allUniqueRecordsMap.containsKey(poNumber)) {
                        String[] existing = allUniqueRecordsMap.get(poNumber);
                        String existingSdn = (existing.length > COL_SALORD_NO) ? existing[COL_SALORD_NO] : "";
                        
                        boolean existingHasSdn = existingSdn != null && !existingSdn.trim().isEmpty(); // Any text/number is valid
                        boolean newHasSdn = sdn != null && !sdn.trim().isEmpty(); // Any text/number is valid

                        // If existing entry has NO SDN, but the new one DOES, replace it.
                        if (!existingHasSdn && newHasSdn) {
                            allUniqueRecordsMap.put(poNumber, rowData);
                        }
                    } else {
                        allUniqueRecordsMap.put(poNumber, rowData);
                    }
                }
            }
        }

        // Separate into lists based on SDN presence
        List<String[]> recordsWithSdn = new ArrayList<>();
        Map<String, String[]> recordsWithoutSdnMap = new LinkedHashMap<>();

        for (Map.Entry<String, String[]> entry : allUniqueRecordsMap.entrySet()) {
            String[] rowData = entry.getValue();
            String sdn = (rowData.length > COL_SALORD_NO) ? rowData[COL_SALORD_NO] : "";
            
            if (sdn != null && !sdn.trim().isEmpty()) { // Consider ANY non-blank text as SDN
                recordsWithSdn.add(rowData);
            } else {
                recordsWithoutSdnMap.put(entry.getKey(), rowData);
            }
        }

        System.out.println("Read Complete.");
        System.out.println("Total Rows Read: " + totalRowsRead);
        System.out.println("Total Unique Records (after removing duplicates): " + allUniqueRecordsMap.size());
        System.out.println("Records with SDN: " + recordsWithSdn.size());
        System.out.println("Unique Records without SDN: " + recordsWithoutSdnMap.size());

        // 2. Process SAP Checks
        SapService sapService = new SapService();
        List<String> posToCheck = new ArrayList<>(recordsWithoutSdnMap.keySet());
        Map<String, Map<String, String>> sapResults = new LinkedHashMap<>();

        if (!posToCheck.isEmpty()) {
            sapResults = sapService.runSapAutomation(posToCheck, sapConnectionName);
        }

        List<String[]> processedWithoutSdn = new ArrayList<>();
        
        // Add Remark header
        headers.add("Remark");
        headers.add("Final Remark");

        // Update recordsWithSdn to include new columns and set "Interfaced"
        List<String[]> updatedRecordsWithSdn = new ArrayList<>();
        for (String[] rowData : recordsWithSdn) {
            if (rowData.length < headers.size()) {
                String[] newRowData = new String[headers.size()];
                System.arraycopy(rowData, 0, newRowData, 0, rowData.length);
                rowData = newRowData;
            }
            // Set Final Remark (last column)
            rowData[headers.size() - 1] = "Interfaced";
            updatedRecordsWithSdn.add(rowData);
        }
        recordsWithSdn = updatedRecordsWithSdn;

        for (Map.Entry<String, String[]> entry : recordsWithoutSdnMap.entrySet()) {
            String po = entry.getKey();
            String[] rowData = entry.getValue();

            // Resize rowData to fit Remark and Final Remark columns if needed
            if (rowData.length < headers.size()) {
                String[] newRowData = new String[headers.size()];
                System.arraycopy(rowData, 0, newRowData, 0, rowData.length);
                rowData = newRowData;
            }

            Map<String, String> sapResult = sapResults.get(po);
            String remark;

            if (sapResult != null) {
                String sapSdn = sapResult.get("SDN");
                String sapMsg = sapResult.get("MESSAGE");

                if (sapSdn != null && !sapSdn.trim().isEmpty()) {
                    rowData[COL_SALORD_NO] = sapSdn;
                    remark = "Added from SAP";
                } else {
                    remark = sapMsg != null ? sapMsg : "No SDN found in SAP";
                }
            } else {
                remark = "File Not Trigger from SAP";
            }

            // Set Remark (second to last column)
            rowData[headers.size() - 2] = remark;

            // Set Final Remark (last column)
            String finalRemark = "";
            if (remark != null && remark.toLowerCase().contains("credit")) {
                finalRemark = "Credit Balance issue";
            } else if (rowData[COL_SALORD_NO] != null && !rowData[COL_SALORD_NO].trim().isEmpty()) {
                finalRemark = "Interfaced";
            } else {
                finalRemark = remark; // Copy other remarks (like errors) to final remark
            }
            rowData[headers.size() - 1] = finalRemark;
            
            processedWithoutSdn.add(rowData);
        }

        // 3. Write Output
        try (SXSSFWorkbook outputWorkbook = new SXSSFWorkbook(100)) {
            // Create Styles
            CellStyle headerStyle = outputWorkbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            Font headerFont = outputWorkbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);

            CellStyle dataStyle = outputWorkbook.createCellStyle();
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderTop(BorderStyle.THIN);
            dataStyle.setBorderLeft(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);
            
            CellStyle totalStyle = outputWorkbook.createCellStyle();
            totalStyle.cloneStyleFrom(dataStyle);
            totalStyle.setFont(headerFont);

            Sheet outSheet = outputWorkbook.createSheet("Processed Data");
            ((SXSSFSheet) outSheet).trackAllColumnsForAutoSizing();

            int rowIndex = 0;

            Row headerRow = outSheet.createRow(rowIndex++);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
                cell.setCellStyle(headerStyle);
            }

            for (String[] data : recordsWithSdn) {
                writeRow(outSheet.createRow(rowIndex++), data, dataStyle);
            }

            for (String[] data : processedWithoutSdn) {
                writeRow(outSheet.createRow(rowIndex++), data, dataStyle);
            }

            for (int i = 0; i < headers.size(); i++) {
                outSheet.autoSizeColumn(i);
            }

            // 4. Create Summary Pivot Table
            System.out.println("Creating summary sheet...");

            // First, calculate the summary counts
            Map<String, Long> summaryCounts = new LinkedHashMap<>();
            int finalRemarkIndex = headers.size() - 1;

            for (String[] rowData : recordsWithSdn) {
                String finalRemark = (rowData.length > finalRemarkIndex && rowData[finalRemarkIndex] != null) ? rowData[finalRemarkIndex] : "Blank";
                summaryCounts.put(finalRemark, summaryCounts.getOrDefault(finalRemark, 0L) + 1);
            }

            for (String[] rowData : processedWithoutSdn) {
                String finalRemark = (rowData.length > finalRemarkIndex && rowData[finalRemarkIndex] != null) ? rowData[finalRemarkIndex] : "Blank";
                summaryCounts.put(finalRemark, summaryCounts.getOrDefault(finalRemark, 0L) + 1);
            }

            // Now, write the summary to a new sheet
            Sheet summarySheet = outputWorkbook.createSheet("Summary");
            ((SXSSFSheet) summarySheet).trackAllColumnsForAutoSizing();
            int summaryRowIndex = 0;

            // Header for summary
            Row summaryHeaderRow = summarySheet.createRow(summaryRowIndex++);
            Cell h1 = summaryHeaderRow.createCell(0); h1.setCellValue("Final Remark"); h1.setCellStyle(headerStyle);
            Cell h2 = summaryHeaderRow.createCell(1); h2.setCellValue("Count of POs"); h2.setCellStyle(headerStyle);

            long totalSummaryCount = 0;
            // Data for summary
            for (Map.Entry<String, Long> summaryEntry : summaryCounts.entrySet()) {
                Row summaryDataRow = summarySheet.createRow(summaryRowIndex++);
                Cell c1 = summaryDataRow.createCell(0); c1.setCellValue(summaryEntry.getKey()); c1.setCellStyle(dataStyle);
                Cell c2 = summaryDataRow.createCell(1); c2.setCellValue(summaryEntry.getValue()); c2.setCellStyle(dataStyle);
                totalSummaryCount += summaryEntry.getValue();
            }
            
            // Total Row for Summary
            Row summaryTotalRow = summarySheet.createRow(summaryRowIndex++);
            Cell t1 = summaryTotalRow.createCell(0); t1.setCellValue("Total"); t1.setCellStyle(totalStyle);
            Cell t2 = summaryTotalRow.createCell(1); t2.setCellValue(totalSummaryCount); t2.setCellStyle(totalStyle);

            summarySheet.autoSizeColumn(0);
            summarySheet.autoSizeColumn(1);

            // 5. Create Spart-based reports
            List<String[]> allFinalRecords = new ArrayList<>();
            allFinalRecords.addAll(recordsWithSdn);
            allFinalRecords.addAll(processedWithoutSdn);
            createSpartReports(outputWorkbook, headers, allFinalRecords, headerStyle, dataStyle, totalStyle);
            createAuartReport(outputWorkbook, headers, allFinalRecords, headerStyle, dataStyle, totalStyle);

            try {
                try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                    outputWorkbook.write(fos);
                }
            } catch (java.io.FileNotFoundException e) {
                System.err.println("ERROR: Could not open output file '" + outputPath + "'. Please close it if it is open in Excel.");
                throw e;
            }
            outputWorkbook.dispose();
        }

        System.out.println("Processing Complete. Output saved to: " + outputPath);
    }

    private void writeRow(Row row, String[] data, CellStyle style) {
        for (int i = 0; i < data.length; i++) {
            if (data[i] != null) {
                Cell cell = row.createCell(i);
                cell.setCellValue(data[i]);
                if (style != null) {
                    cell.setCellStyle(style);
                }
            }
        }
    }

    private void createSpartReports(SXSSFWorkbook workbook, List<String> headers, List<String[]> allFinalRecords, CellStyle headerStyle, CellStyle dataStyle, CellStyle totalStyle) {
        System.out.println("Creating Spart-based summary reports...");

        Map<String, String> spartMap = new LinkedHashMap<>();
        spartMap.put("02", "Vehicle");
        spartMap.put("03", "Spares");
        spartMap.put("08", "Oil");
        spartMap.put("04", "Gear");
        spartMap.put("10", "GMA");
        // Add single digit variants for robustness
        

        // Initialize data structures for summaries
        Map<String, SpartSummaryData> spartSummary = new LinkedHashMap<>();
        for (String name : spartMap.values()) {
            spartSummary.put(name, new SpartSummaryData());
        }

        Map<String, Map<String, Long>> issueSummary = new LinkedHashMap<>();
        for (String name : spartMap.values()) {
            issueSummary.put(name, new LinkedHashMap<>());
        }

        // Populate summary data
        int finalRemarkIndex = headers.size() - 1;
        for (String[] rowData : allFinalRecords) {
            String spartCode = (rowData.length > COL_SPART && rowData[COL_SPART] != null) ? rowData[COL_SPART].trim() : "";
            
            // Handle numeric formatting (e.g., "2.0" -> "2")
            if (spartCode.endsWith(".0")) {
                spartCode = spartCode.substring(0, spartCode.length() - 2);
            }

            String spartName = spartMap.get(spartCode);
            if (spartName == null && spartCode.length() == 1) {
                 spartName = spartMap.get("0" + spartCode);
            }

            if (spartName != null) {
                SpartSummaryData summaryData = spartSummary.get(spartName);
                summaryData.totalCases++;

                String finalRemark = (rowData.length > finalRemarkIndex && rowData[finalRemarkIndex] != null) ? rowData[finalRemarkIndex] : "Blank";

                if ("Interfaced".equalsIgnoreCase(finalRemark)) {
                    summaryData.interfacedCount++;
                } else {
                    Map<String, Long> spartIssues = issueSummary.get(spartName);
                    spartIssues.put(finalRemark, spartIssues.getOrDefault(finalRemark, 0L) + 1);
                }
            } else if (!spartCode.isEmpty()) {
                System.out.println("WARNING: Unmapped Spart code found: '" + spartCode + "'");
            }
        }

        // --- Create "Spart Summary" Sheet ---
        Sheet spartSummarySheet = workbook.createSheet("Spart Summary");
        ((SXSSFSheet) spartSummarySheet).trackAllColumnsForAutoSizing();
        int rowIndex = 0;
        Row headerRow = spartSummarySheet.createRow(rowIndex++);
        String[] spartHeaders = {"Spart", "Total Cases", "Interfaced", "Pending for Interface", "Success Percentage"};
        for (int i = 0; i < spartHeaders.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(spartHeaders[i]);
            cell.setCellStyle(headerStyle);
        }

        long grandTotalCases = 0, grandTotalInterfaced = 0;
        NumberFormat percentFormat = NumberFormat.getPercentInstance();
        percentFormat.setMinimumFractionDigits(1);

        for (Map.Entry<String, SpartSummaryData> entry : spartSummary.entrySet()) {
            Row dataRow = spartSummarySheet.createRow(rowIndex++);
            SpartSummaryData data = entry.getValue();
            long pending = data.totalCases - data.interfacedCount;
            double successRate = (data.totalCases > 0) ? (double) data.interfacedCount / data.totalCases : 0;

            Cell c0 = dataRow.createCell(0); c0.setCellValue(entry.getKey()); c0.setCellStyle(dataStyle);
            Cell c1 = dataRow.createCell(1); c1.setCellValue(data.totalCases); c1.setCellStyle(dataStyle);
            Cell c2 = dataRow.createCell(2); c2.setCellValue(data.interfacedCount); c2.setCellStyle(dataStyle);
            Cell c3 = dataRow.createCell(3); c3.setCellValue(pending); c3.setCellStyle(dataStyle);
            Cell c4 = dataRow.createCell(4); c4.setCellValue(percentFormat.format(successRate)); c4.setCellStyle(dataStyle);

            grandTotalCases += data.totalCases;
            grandTotalInterfaced += data.interfacedCount;
        }

        // Total Row
        Row totalRow = spartSummarySheet.createRow(rowIndex++);
        long grandTotalPending = grandTotalCases - grandTotalInterfaced;
        double grandSuccessRate = (grandTotalCases > 0) ? (double) grandTotalInterfaced / grandTotalCases : 0;
        Cell t0 = totalRow.createCell(0); t0.setCellValue("Total"); t0.setCellStyle(totalStyle);
        Cell t1 = totalRow.createCell(1); t1.setCellValue(grandTotalCases); t1.setCellStyle(totalStyle);
        Cell t2 = totalRow.createCell(2); t2.setCellValue(grandTotalInterfaced); t2.setCellStyle(totalStyle);
        Cell t3 = totalRow.createCell(3); t3.setCellValue(grandTotalPending); t3.setCellStyle(totalStyle);
        Cell t4 = totalRow.createCell(4); t4.setCellValue(percentFormat.format(grandSuccessRate)); t4.setCellStyle(totalStyle);

        for (int i = 0; i < spartHeaders.length; i++) {
            spartSummarySheet.autoSizeColumn(i);
        }

        // --- Create "Issue Summary" Sheet ---
        Sheet issueSummarySheet = workbook.createSheet("Issue Summary");
        ((SXSSFSheet) issueSummarySheet).trackAllColumnsForAutoSizing();
        rowIndex = 0;
        headerRow = issueSummarySheet.createRow(rowIndex++);
        Cell h0 = headerRow.createCell(0); h0.setCellValue("Issue"); h0.setCellStyle(headerStyle);
        List<String> spartNames = new ArrayList<>(spartMap.values());
        for (int i = 0; i < spartNames.size(); i++) {
            Cell cell = headerRow.createCell(i + 1);
            cell.setCellValue(spartNames.get(i));
            cell.setCellStyle(headerStyle);
        }
        Cell hTotal = headerRow.createCell(spartNames.size() + 1); hTotal.setCellValue("Total"); hTotal.setCellStyle(headerStyle);

        // Collect all unique issues
        issueSummary.values().stream().flatMap(m -> m.keySet().stream()).distinct().forEach(issue -> {
            Row dataRow = issueSummarySheet.createRow(issueSummarySheet.getLastRowNum() + 1);
            Cell cIssue = dataRow.createCell(0); cIssue.setCellValue(issue); cIssue.setCellStyle(dataStyle);
            long rowTotal = 0;
            for (int i = 0; i < spartNames.size(); i++) {
                String spartName = spartNames.get(i);
                long count = issueSummary.get(spartName).getOrDefault(issue, 0L);
                Cell cCount = dataRow.createCell(i + 1); cCount.setCellValue(count); cCount.setCellStyle(dataStyle);
                rowTotal += count;
            }
            Cell cRowTotal = dataRow.createCell(spartNames.size() + 1); cRowTotal.setCellValue(rowTotal); cRowTotal.setCellStyle(dataStyle);
        });

        for (int i = 0; i <= spartNames.size() + 1; i++) {
            issueSummarySheet.autoSizeColumn(i);
        }
    }

    private void createAuartReport(SXSSFWorkbook workbook, List<String> headers, List<String[]> allFinalRecords, CellStyle headerStyle, CellStyle dataStyle, CellStyle totalStyle) {
        System.out.println("Creating Auart-based summary report...");

        int auartColIndex = -1;
        for (int i = 0; i < headers.size(); i++) {
            if ("AUART".equalsIgnoreCase(headers.get(i))) {
                auartColIndex = i;
                break;
            }
        }

        if (auartColIndex == -1) {
            System.out.println("WARNING: 'AUART' column not found in input file. Skipping 'Auart Remarks' sheet.");
            return;
        }

        // Data structure: Item -> SubHeading -> SummaryData
        Map<String, Map<String, SpartSummaryData>> auartSummary = new LinkedHashMap<>();

        // Spart mapping for ZUB resolution
        Map<String, String> spartMap = new LinkedHashMap<>();
        spartMap.put("02", "Vehicle");
        spartMap.put("03", "Spares");
        spartMap.put("08", "Oil");
        spartMap.put("04", "Gear");
        spartMap.put("10", "GMA");

        // Populate summary data
        int finalRemarkIndex = headers.size() - 1;
        for (String[] rowData : allFinalRecords) {
            String auartCode = (rowData.length > auartColIndex && rowData[auartColIndex] != null) ? rowData[auartColIndex].trim() : "";
            String spartCode = (rowData.length > COL_SPART && rowData[COL_SPART] != null) ? rowData[COL_SPART].trim() : "";

            if (spartCode.endsWith(".0")) {
                spartCode = spartCode.substring(0, spartCode.length() - 2);
            }
            String spartName = spartMap.get(spartCode);
            if (spartName == null && spartCode.length() == 1) {
                spartName = spartMap.get("0" + spartCode);
            }

            AuartInfo info = getAuartInfo(auartCode, spartName);

            if (info != null) {
                Map<String, SpartSummaryData> subHeadingMap = auartSummary.computeIfAbsent(info.item, k -> new LinkedHashMap<>());
                SpartSummaryData summaryData = subHeadingMap.computeIfAbsent(info.subHeading, k -> new SpartSummaryData());

                summaryData.totalCases++;
                String finalRemark = (rowData.length > finalRemarkIndex && rowData[finalRemarkIndex] != null) ? rowData[finalRemarkIndex] : "Blank";
                if ("Interfaced".equalsIgnoreCase(finalRemark)) {
                    summaryData.interfacedCount++;
                }
            }
        }

        // --- Create "Auart Remarks" Sheet ---
        Sheet sheet = workbook.createSheet("Auart Remarks");
        ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        int rowIndex = 0;

        CellStyle itemHeaderStyle = workbook.createCellStyle();
        itemHeaderStyle.cloneStyleFrom(headerStyle);
        itemHeaderStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());

        NumberFormat percentFormat = NumberFormat.getPercentInstance();
        percentFormat.setMinimumFractionDigits(1);

        String[] subHeaders = {"Sub-Heading", "Total Cases", "Interfaced", "Pending for Interface", "Success Percentage"};

        for (Map.Entry<String, Map<String, SpartSummaryData>> itemEntry : auartSummary.entrySet()) {
            String item = itemEntry.getKey();
            Map<String, SpartSummaryData> subHeadingMap = itemEntry.getValue();

            Row itemHeaderRow = sheet.createRow(rowIndex++);
            for (int i = 0; i < subHeaders.length; i++) {
                Cell cell = itemHeaderRow.createCell(i);
                if (i == 0) cell.setCellValue(item);
                cell.setCellStyle(itemHeaderStyle);
            }
            sheet.addMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, subHeaders.length - 1));

            Row headerRow = sheet.createRow(rowIndex++);
            for (int i = 0; i < subHeaders.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(subHeaders[i]);
                cell.setCellStyle(headerStyle);
            }

            long itemTotalCases = 0, itemTotalInterfaced = 0;

            for (Map.Entry<String, SpartSummaryData> subHeadingEntry : subHeadingMap.entrySet()) {
                Row dataRow = sheet.createRow(rowIndex++);
                SpartSummaryData data = subHeadingEntry.getValue();
                long pending = data.totalCases - data.interfacedCount;
                double successRate = (data.totalCases > 0) ? (double) data.interfacedCount / data.totalCases : 0;

                Cell c0 = dataRow.createCell(0); c0.setCellValue(subHeadingEntry.getKey()); c0.setCellStyle(dataStyle);
                Cell c1 = dataRow.createCell(1); c1.setCellValue(data.totalCases); c1.setCellStyle(dataStyle);
                Cell c2 = dataRow.createCell(2); c2.setCellValue(data.interfacedCount); c2.setCellStyle(dataStyle);
                Cell c3 = dataRow.createCell(3); c3.setCellValue(pending); c3.setCellStyle(dataStyle);
                Cell c4 = dataRow.createCell(4); c4.setCellValue(percentFormat.format(successRate)); c4.setCellStyle(dataStyle);
            }
            rowIndex++; // Add a blank row for spacing
        }

        for (int i = 0; i < subHeaders.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    // --- Inner Classes ---

    private static class SpartSummaryData {
        long totalCases = 0;
        long interfacedCount = 0;
    }

    public static class SapService {
        
        public Map<String, Map<String, String>> runSapAutomation(List<String> poNumbers, String connectionName) {
            System.out.println("Connecting to SAP using connection name: " + connectionName);
            Map<String, Map<String, String>> results = new LinkedHashMap<>();
            String user = "Bijus";
            String pass = "Royal@26";

            try {
                // 1. Copy POs to Clipboard for SAP Paste
                String clipboardData = String.join("\r\n", poNumbers);
                StringSelection selection = new StringSelection(clipboardData);
                Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                clipboard.setContents(selection, selection);
                System.out.println("Copied " + poNumbers.size() + " PO numbers to clipboard.");

                // Ensure SAP Logon is running
                Process checkSap = new ProcessBuilder("tasklist", "/FI", "IMAGENAME eq saplogon.exe").start();
                try (Scanner scanner = new Scanner(checkSap.getInputStream())) {
                    if (!scanner.useDelimiter("\\A").next().contains("saplogon.exe")) {
                        System.out.println("Starting SAP Logon...");
                        new ProcessBuilder("C:\\Program Files\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe").start();
                        Thread.sleep(5000); // Wait for SAP to open
                    }
                }

                // Prepare Temp File for Export
                String tempDir = System.getProperty("java.io.tmpdir");
                if (!tempDir.endsWith(File.separator)) tempDir += File.separator;
                String exportFileName = "sap_export.txt";
                File exportFile = new File(tempDir + exportFileName);
                if (exportFile.exists()) exportFile.delete();
                
                String debugLogFileName = "sap_debug_log.txt";
                File debugLogFile = new File(tempDir + debugLogFileName);
                if (debugLogFile.exists()) debugLogFile.delete();
                System.out.println("SAP Debug Log will be written to: " + debugLogFile.getAbsolutePath());

                StringBuilder vbsContent = new StringBuilder();
                // --- VBScript Generation ---
                vbsContent.append("If Not IsObject(application) Then\n");
                vbsContent.append("   Set SapGuiAuto  = GetObject(\"SAPGUI\")\n");
                vbsContent.append("   Set application = SapGuiAuto.GetScriptingEngine\n");
                vbsContent.append("End If\n");
                
                // Debug Logging setup
                vbsContent.append("Set fso = CreateObject(\"Scripting.FileSystemObject\")\n");
                vbsContent.append("Set debugLog = fso.CreateTextFile(\"").append(debugLogFile.getAbsolutePath().replace("\\", "\\\\")).append("\", True)\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] VBScript Started.\"\n");
                vbsContent.append("On Error Resume Next\n"); // Enable error handling for VBScript

                vbsContent.append("Set connection = application.OpenConnection(\"").append(connectionName).append("\", True)\n");
                vbsContent.append("If Err.Number <> 0 Then debugLog.WriteLine \"[\" & Now & \"] ERROR: Could not open connection: \" & Err.Description & \" (\" & Err.Number & \")\"\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Connected to SAP: \" & (Not connection Is Nothing)\n");
                vbsContent.append("Set session = connection.Children(0)\n");
                // Force window to front using WScript.Shell
                vbsContent.append("WScript.Sleep 500\n");
                vbsContent.append("Set WshShell = CreateObject(\"WScript.Shell\")\n");
                vbsContent.append("WshShell.AppActivate \"SAP\"\n");
                // Maximize SAP window to ensure it is visible
                vbsContent.append("session.findById(\"wnd[0]\").maximize\n");
                vbsContent.append("session.findById(\"wnd[0]\").setFocus\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] SAP window maximized and focused.\"\n");

                // Login logic
                vbsContent.append("session.findById(\"wnd[0]/usr/txtRSYST-BNAME\").text = \"").append(user).append("\"\n");
                vbsContent.append("session.findById(\"wnd[0]/usr/pwdRSYST-BCODE\").text = \"").append(pass).append("\"\n");
                vbsContent.append("session.findById(\"wnd[0]\").sendVKey 0\n");
                vbsContent.append("If Err.Number <> 0 Then debugLog.WriteLine \"[\" & Now & \"] ERROR: Login failed: \" & Err.Description & \" (\" & Err.Number & \")\"\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Login attempt complete.\"\n");

                vbsContent.append("session.findById(\"wnd[0]/tbar[0]/okcd\").text = \"/nSE16\"\n");
                vbsContent.append("session.findById(\"wnd[0]\").sendVKey 0\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Navigated to SE16.\"\n");
                vbsContent.append("session.findById(\"wnd[0]/usr/ctxtDATABROWSE-TABLENAME\").text = \"ZSD_CDS_PO_IF\"\n");
                vbsContent.append("session.findById(\"wnd[0]\").sendVKey 0\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Entered table ZSD_CDS_PO_IF.\"\n");

                // Click 'Multiple Selection' button for the POR field
                vbsContent.append("session.findById(\"wnd[0]/usr/btn%_I4_%_APP_%-VALU_PUSH\").press\n");
                vbsContent.append("WScript.Sleep 2000\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Clicked Multiple Selection button.\"\n");
                
                // Paste from Clipboard (Button ID 24 is standard for 'Upload from Clipboard')
                vbsContent.append("session.findById(\"wnd[1]/tbar[0]/btn[24]\").press\n");
                vbsContent.append("WScript.Sleep 8000\n"); // Increased wait for large data paste

                vbsContent.append("session.findById(\"wnd[1]/tbar[0]/btn[8]\").press\n"); // Copy/Execute selection
                vbsContent.append("WScript.Sleep 2000\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Pasted POs from clipboard and executed selection.\"\n");
                
                // Execute Query
                vbsContent.append("session.findById(\"wnd[0]/tbar[1]/btn[8]\").press\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Executed main query.\"\n");
                
                // --- New Grid Interaction Logic ---
                vbsContent.append("On Error Resume Next\n");
                
                // Wait for Grid Object to appear (handle screen transition delay)
                vbsContent.append("Dim grid\n");
                vbsContent.append("Set grid = Nothing\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] Starting wait for grid object (up to 60s).\"\n");
                vbsContent.append("For k = 1 To 60\n"); // Wait up to 60 seconds for result screen
                vbsContent.append("  Err.Clear\n");
                vbsContent.append("  Set grid = session.findById(\"wnd[0]/usr/cntlGRID1/shellcont/shell\")\n");
                vbsContent.append("  If Err.Number = 0 Then Exit For\n");
                vbsContent.append("  WScript.Sleep 1000\n");
                vbsContent.append("Next\n");
                vbsContent.append("If Not grid Is Nothing Then\n");
                vbsContent.append("  debugLog.WriteLine \"[\" & Now & \"] Grid object found.\"\n");
                vbsContent.append("Else\n");
                vbsContent.append("  debugLog.WriteLine \"[\" & Now & \"] ERROR: Grid object NOT found after 60s. Last error: \" & Err.Description & \" (\" & Err.Number & \")\"\n");
                vbsContent.append("End If\n");

                vbsContent.append("If Not grid Is Nothing Then\n");
                
                // Add a dynamic wait loop for the grid to be populated with data
                vbsContent.append("  debugLog.WriteLine \"[\" & Now & \"] Starting wait for grid data (up to 120s).\"\n");
                vbsContent.append("  For j = 1 to 120\n"); // Wait up to 120 seconds for data to load
                vbsContent.append("    If grid.RowCount > 0 Then Exit For\n");
                vbsContent.append("    WScript.Sleep 1000\n");
                vbsContent.append("  Next\n");
                vbsContent.append("  debugLog.WriteLine \"[\" & Now & \"] Grid data check complete. RowCount: \" & grid.RowCount & \".\"\n");

                // Only proceed if rows were actually found
                vbsContent.append("  If grid.RowCount > 0 Then\n");
                
                // Sort by Document (SDN) column descending
                vbsContent.append("    grid.setCurrentCell -1, \"SDN\"\n");
                vbsContent.append("    grid.selectColumn \"SDN\"\n");
                vbsContent.append("    session.findById(\"wnd[0]/tbar[1]/btn[40]\").press\n"); // Descending sort button
                vbsContent.append("    WScript.Sleep 2000\n");
                vbsContent.append("    debugLog.WriteLine \"[\" & Now & \"] Grid sorted by SDN descending.\"\n");
                
                // Re-acquire grid object after sort to ensure validity
                vbsContent.append("    Set grid = session.findById(\"wnd[0]/usr/cntlGRID1/shellcont/shell\")\n");

                // Write to file using FileSystemObject
                vbsContent.append("    Set fso = CreateObject(\"Scripting.FileSystemObject\")\n");
                vbsContent.append("    Set f = fso.CreateTextFile(\"").append(exportFile.getAbsolutePath().replace("\\", "\\\\")).append("\", True)\n");
                vbsContent.append("    rowCount = grid.RowCount\n");
                vbsContent.append("    debugLog.WriteLine \"[\" & Now & \"] Starting data extraction loop for \" & rowCount & \" rows.\"\n");
                vbsContent.append("    For i = 0 To rowCount - 1\n");
                vbsContent.append("      On Error Resume Next\n");
                // Ensure the row is selected/visible before reading (use setCurrentCell)
                // vbsContent.append("      grid.setCurrentCell i, \"DMSPONO\"\n");
                // vbsContent.append("      WScript.Sleep 300\n");
                vbsContent.append("      If i Mod 20 = 0 Then\n");
                vbsContent.append("        grid.firstVisibleRow = i\n");
                vbsContent.append("        WScript.Sleep 200\n");
                vbsContent.append("      End If\n");
                vbsContent.append("      por_val = \"\"\n");
                vbsContent.append("      sdn_val = \"\"\n");
                vbsContent.append("      msg_val = \"\"\n");
                vbsContent.append("      readAttempts = 0\n");
                vbsContent.append("      Do While readAttempts < 5\n");
                vbsContent.append("        On Error Resume Next\n");
                vbsContent.append("        por_val = grid.getCellValue(i, \"DMSPONO\")\n");
                vbsContent.append("        sdn_val = grid.getCellValue(i, \"SDN\")\n");
                vbsContent.append("        msg_val = grid.getCellValue(i, \"MESSAGE\")\n");
                vbsContent.append("        If Err.Number = 0 And Trim(por_val) <> \"\" Then Exit Do\n");
                vbsContent.append("        readAttempts = readAttempts + 1\n");
                vbsContent.append("        Err.Clear\n");
                vbsContent.append("        WScript.Sleep 200\n");
                vbsContent.append("      Loop\n");
                vbsContent.append("      If Err.Number = 0 Then\n");
                vbsContent.append("        f.WriteLine por_val & vbTab & sdn_val & vbTab & msg_val\n");
                vbsContent.append("      Else\n");
                vbsContent.append("        debugLog.WriteLine \"[\" & Now & \"] ERROR: Reading row \" & i & \": \" & Err.Description & \" (\" & Err.Number & \")\"\n");
                vbsContent.append("        Err.Clear\n");
                vbsContent.append("      End If\n");
                vbsContent.append("    Next\n");
                vbsContent.append("    f.Close\n");
                vbsContent.append("    debugLog.WriteLine \"[\" & Now & \"] Data extraction loop complete. Export file written.\"\n");
                vbsContent.append("  End If\n"); // End check for RowCount > 0
                vbsContent.append("End If\n");
                vbsContent.append("On Error GoTo 0\n");

                // Close SAP Window
                vbsContent.append("session.findById(\"wnd[0]/tbar[0]/okcd\").text = \"/nex\"\n");
                vbsContent.append("session.findById(\"wnd[0]\").sendVKey 0\n");
                vbsContent.append("debugLog.WriteLine \"[\" & Now & \"] VBScript Finished.\"\n");
                vbsContent.append("debugLog.Close\n"); // Close the debug log file

                Path scriptPath = Files.createTempFile("sap_script", ".vbs");
                Files.writeString(scriptPath, vbsContent.toString());
                Process p = new ProcessBuilder("wscript", scriptPath.toAbsolutePath().toString()).start();
                p.waitFor();

                // Wait for file to appear (polling)
                for (int i = 0; i < 15; i++) {
                    if (exportFile.exists() && exportFile.length() > 0) break;
                    Thread.sleep(1000);
                }

                // Read the exported file
                if (exportFile.exists()) {
                    try (BufferedReader br = new BufferedReader(new FileReader(exportFile, java.nio.charset.Charset.defaultCharset()))) {
                        String line;
                        
                        while ((line = br.readLine()) != null) {
                            // The new VBScript writes a simple "PO \t SDN \t MESSAGE" format.
                            String[] parts = line.split("\t", -1); // Use -1 to keep trailing empty parts
                            if (parts.length >= 1) {
                                String po = parts[0].trim();
                                if (po.isEmpty()) continue; // Skip lines that don't have a PO number

                                String newSdn = (parts.length > 1) ? parts[1].trim() : "";
                                String newMsg = (parts.length > 2) ? parts[2].trim() : "";

                                // If we haven't seen this PO before, add it to our results.
                                if (!results.containsKey(po)) {
                                    Map<String, String> data = new LinkedHashMap<>();
                                    data.put("SDN", newSdn);
                                    data.put("MESSAGE", newMsg);
                                    results.put(po, data);
                                } else {
                                    // An entry for this PO already exists. Check if we should update it.
                                    Map<String, String> existingData = results.get(po);
                                    String existingSdn = existingData.getOrDefault("SDN", "");

                                    // If the existing record has no SDN, but this new line does, update the record.
                                    if (existingSdn.isEmpty() && !newSdn.isEmpty()) {
                                        existingData.put("SDN", newSdn);
                                        existingData.put("MESSAGE", newMsg); // Also update the message
                                    }
                                }
                            }
                        }
                        System.out.println("SAP Extraction Complete. Found data for " + results.size() + " POs.");
                    }
                } else {
                    System.err.println("SAP Export file was not created. Check SAP execution.");
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
            return results;
        }
    }
}
