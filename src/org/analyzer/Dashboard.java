package org.analyzer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.print.Book;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.*;
import java.util.List;

/**
 * Created by vhilc on 12/11/2016.
 */
public class Dashboard {

    public JPanel panel1;
    private JTextField directoryField;
    private JButton createReportButton;
    private JTextField columnReport;
    private Map<String, List<List<String>>> recordsFromExcel = new HashMap<>();

    public Dashboard() {

        createReportButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String directory = directoryField.getText();
                String columnName = columnReport.getText();
                if (directory == null || columnName == null || directory.length() == 0 || columnName.length() == 0) {
                    displayError("Грешка! Някое от полетата е празно!");
                } else {
                    buildReport(directory, columnName);
                }

            }
        });
    }

    private void buildReport(String directory, String columnName) throws IOException {
        System.out.println(directory + "\n" + columnName);
        File folder = new File(directory);
        File[] listOfFiles = folder.listFiles();
        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile()) {
                    System.out.println("File " + file.getName());
                    readRecordsFromExcel(columnName, file.getName());
                }
            }
        } else {
            displayError("Няма намерени ексел файлове в директория:" + directory + "! Грешна директория?");
        }
    }

    public Map<String, List<List<String>>> readRecordsFromExcel(String columnName, String excelFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

        Workbook workbook;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            //not supported excel
            return null;
        }
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();
        int numMergedRegions = firstSheet.getNumMergedRegions();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            //skip merged cells in begin of file
            if (nextRow.getRowNum() <= numMergedRegions) {
                continue;
            }
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            String columnValue;
            int i = 0;
            while (cellIterator.hasNext()) {
                i++;
                Cell nextCell = cellIterator.next();
                int columnIndex = nextCell.getColumnIndex();
                if (columnIndex == cellReportIndex) {
                    columnValue = getCellValue((String)nextCell;
                }

                switch (columnIndex) {
                    case 1:
                        aBook.setTitle((String) getCellValue(nextCell));
                        break;
                    case 2:
                        aBook.setAuthor((String) getCellValue(nextCell));
                        break;
                    case 3:
                        aBook.setPrice((double) getCellValue(nextCell));

                        break;
                }


            }
            readRecordsFromExcel().put(cellName, )
        }

        workbook.close();
        inputStream.close();

        return listBooks;
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

    private void displayError(String errorMessage) {
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
