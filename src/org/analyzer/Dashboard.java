package org.analyzer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.util.*;
import java.util.List;

public class Dashboard {

    public JPanel panel1;
    private JTextField directoryField;
    private JButton createReportButton;
    private JTextField columnReport;
    private Map<String, List<RecordEntity>> recordsFromExcel;
    private StringBuffer reportBuffer;
    private static final String SEPARATOR = "================================";

    public Dashboard() {

        createReportButton.addActionListener(e -> {
            recordsFromExcel = new HashMap<>();
            reportBuffer = new StringBuffer();
            appendToBuffer(Calendar.getInstance().getTime().toString());
            String directory = directoryField.getText();
            String columnName = columnReport.getText();
            if (directory == null || columnName == null || directory.length() == 0 || columnName.length() == 0) {
                displayMessage("Грешка! Някое от полетата е празно!");
            } else {
                try {
                    buildReport(directory, columnName);
                    writeReport(directory);
                } catch (IOException e1) {
                    displayMessage("Грешка при обработването на excel:" + e1.getMessage());
                }
            }
        });
    }

    private void writeReport(String directory) {
        Collection<List<RecordEntity>> collection = recordsFromExcel.values();
        List<List<RecordEntity>> collectionList = new ArrayList<>(collection);
        Collections.sort(collectionList, (o1, o2) -> {
            if (o1.size() > o2.size()) {
                return -1;
            } else if (o1.size() < o2.size()) {
                return 1;
            }else  {
                return 0;
            }
        });
        appendToBuffer(SEPARATOR);
        for(List<RecordEntity> recordEntities : collectionList) {
            int recordSize = recordEntities.size();
            appendToBuffer("--");
            appendToBuffer(recordEntities.get(0).getColumnReport() + " срещан " + recordSize + " пъти");
            for (RecordEntity recordEntity : recordEntities) {
                appendToBuffer(recordEntity.toString());
            }
        }

        try{
            String fileTowRite = directory +"\\report.txt";
            PrintWriter writer = new PrintWriter(fileTowRite, "UTF-8");
            writer.println(reportBuffer.toString());
            writer.flush();
            writer.close();
            displayMessage("Създаден успешно репорт: " + fileTowRite);
        } catch (IOException e) {
            // do something
        }
    }

    private void appendToBuffer(String s) {
        reportBuffer.append(s).append(System.getProperty("line.separator"));
    }

    private void buildReport(String directory, String columnName) throws IOException {
        appendToBuffer("Сканиране на директория: " + directory);
        appendToBuffer("Колона за създаване на репорта: " + columnName);
        appendToBuffer(SEPARATOR);
        File folder = new File(directory);
        File[] listOfFiles = folder.listFiles();
        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile()) {
                    readRecordsFromExcel(columnName, file);
                }
            }
        } else {
            displayMessage("Няма намерени ексел файлове в директория:" + directory + "! Грешна директория?");
        }
    }

    public void readRecordsFromExcel(String columnReportName, File excelFile) throws IOException {
        FileInputStream inputStream = new FileInputStream(excelFile);

        Workbook workbook;
        if (excelFile.getName().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFile.getName().endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            //not supported excel
            appendToBuffer("Пропускане на " + excelFile + " тъй като не е Excel file.");
            return;
        }

        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();
        int numMergedRegions = firstSheet.getNumMergedRegions();
        List<String> columnNames = getColumnNames(iterator, numMergedRegions);

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            RecordEntity record = new RecordEntity();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                int columnIndex = currentCell.getColumnIndex();
                String columnName = columnNames.get(columnIndex);
                String cellValue = String.valueOf(getCellValue(currentCell));
                if (columnName.equals(columnReportName)) {
                    record.setColumnReport(cellValue);
                } else {
                    record.put(columnName, cellValue);
                }
            }
            List<RecordEntity> recordEntities = recordsFromExcel.get(record.getColumnReport());
            if (recordEntities == null) {//no previous entries
                recordEntities = new ArrayList();
                recordEntities.add(record);
                recordsFromExcel.put(record.getColumnReport(), recordEntities);
            } else {//add new entry to the existing list
                recordEntities.add(record);
            }
        }
        appendToBuffer("Успешно сканиран " + excelFile);
        workbook.close();
        inputStream.close();
    }

    private List<String> getColumnNames(Iterator<Row> iterator, int numMergedRegions) {
        List<String> columnNames = new ArrayList<>();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            //skip merged cells in begin of file
            if (nextRow.getRowNum() < numMergedRegions) {
                continue;
            }
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                columnNames.add((String) getCellValue(cell));
            }
            break;
        }
        return columnNames;
    }

    private Object getCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return cell.getStringCellValue();

            case BOOLEAN:
                return cell.getBooleanCellValue();

            case NUMERIC:
                return cell.getNumericCellValue();
        }

        return null;
    }

    private void displayMessage(String errorMessage) {
        JOptionPane.showMessageDialog(null, errorMessage);
    }

    private String stringify(String myString) {
        String value = myString;
        try {
            byte bytes[] = myString.getBytes("Windows-1251");
            value = new String(bytes, "UTF-8");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return value;
    }

}
