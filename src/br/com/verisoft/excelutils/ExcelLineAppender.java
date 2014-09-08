package br.com.verisoft.excelutils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLineAppender {

    public static void main(String[] args) {
        long start = System.currentTimeMillis();
        if (args.length < 3) {
            System.out.println("Usage: java -jar xlsx-append.jar [dest-file].xlsx [content] [separator] [line-separator (optional, default \\n)]");
        } else {
            try {
                String lineSeparator = args.length > 3 ? args[3] : "\\\n";
                System.out.println("Starting... DEST-File: " + args[0] + " | Separator: '" + args[2] + " | Line Separator: " + lineSeparator);
                String[] data = args[1].split(lineSeparator);
                System.out.println(Arrays.asList(data));
                Workbook workbook;
                Sheet sheet;
                File file = new File(args[0]);
                if (file.exists()) {
                    InputStream inputStream = new FileInputStream(file);
                    workbook = WorkbookFactory.create(inputStream);
                    sheet = workbook.getSheetAt(0);
                } else {
                    workbook = new XSSFWorkbook();
                    sheet = workbook.createSheet();
                }
                CreationHelper creationHelper = workbook.getCreationHelper();
                int rowNum = sheet.getLastRowNum();
                for (int rowIndex = rowNum; rowIndex < rowNum + data.length; rowIndex++) {
                    String[] csvLineContent = data[rowIndex - rowNum].split(args[2]);
                    Row row = sheet.createRow(rowIndex);
                    if (rowIndex == 0) {
                        for (int columnNumber = 0; columnNumber < sheet.getRow(0).getPhysicalNumberOfCells(); columnNumber++) {
                            sheet.autoSizeColumn(columnNumber);
                        }
                    }
                    for (int cellNumber = 0; cellNumber < csvLineContent.length; cellNumber++) {
                        Cell cell = row.createCell(cellNumber);
                        cell.setCellValue(creationHelper.createRichTextString(csvLineContent[cellNumber]));
                    }

                }
                for (int columnNumber = 0; columnNumber < sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getPhysicalNumberOfCells(); columnNumber++) {
                    int columnWidth = sheet.getColumnWidth(columnNumber);
                    sheet.autoSizeColumn(columnNumber);
                    if (sheet.getColumnWidth(columnNumber) < columnWidth) {
                        sheet.setColumnWidth(columnNumber, columnWidth);
                    }
                }
                FileOutputStream fileOutputStream = new FileOutputStream(args[0]);
                workbook.write(fileOutputStream);
                fileOutputStream.close();
            } catch (Exception exception) {
                exception.printStackTrace();
            }
        }
        System.out.println("Conversion finished! Processing time: " + (System.currentTimeMillis() - start) + "ms.");
    }

}
