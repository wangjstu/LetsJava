package com.wangjstu.NovixysBlog;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

public class PoiExample01 {
    public static void main(String[] args) throws IOException, ParseException {
        System.out.println("Apache POI Excel Example:https://www.novixys.com/blog/apache-poi-excel-example/");
        helloPoiExcel();
    }
    
    public static void helloPoiExcel() throws IOException, ParseException {
        /*The Basics file*/
        // Excel 2007 XLSX file
        Workbook wb = new XSSFWorkbook();
        //older Excel 97-2003 XLS format
        //HSSFWorkbook hssfWorkbook = new HSSFWorkbook();


        /*The First Spreadsheet*/
        //使用WorkbookUtil.createSafeSheetName()，因为某些字符在工作表名称中无效。此实用程序方法将这些字符替换为空格字符。
        String safeSheetName = WorkbookUtil.createSafeSheetName("hello#1");
        Sheet wbSheet = wb.createSheet(safeSheetName);


        /*5. Creating a Row and Cells*/
        Row row = wbSheet.createRow(0);
        row.createCell(0).setCellValue("姓名");
        row.createCell(1).setCellValue("文本");
        row.createCell(2).setCellValue("千分位数字");
        row.createCell(3).setCellValue("Double字段");
        row.createCell(4).setCellValue("数字文本混合");
        row.createCell(5).setCellValue("百分比");

        /*6. Auto Sizing Columns*/
        /*for (int i = wbSheet.getRow(0).getFirstCellNum(), end = wbSheet.getRow(0).getLastCellNum(); i<end;i++) {
            wbSheet.autoSizeColumn(i);
        }*/

        /*7. Multi-Line Column*/
        row.getCell(1).setCellValue("Total\r\n(persons)");
        CellStyle cellStyle = wbSheet.getWorkbook().createCellStyle();
        cellStyle.setWrapText(true);
        row.getCell(1).setCellStyle(cellStyle);

        /*Adding rows*/
        List<List<String>> listVal = Arrays.asList(
                Arrays.asList("wang军", "China","1,378,020,000","147.75","2016","68.12%"),
                Arrays.asList("w军", "United States of America","323,128,000","35.32","2016Q","0%"),
                Arrays.asList("L军", "Indonesia","257,453,000.12","142.12","2016","--"),
                Arrays.asList("M军", "Brazil","206,081,000","24.66","2016","0.02%"),
                Arrays.asList("M军", "Brazil","206,081,000","24.66","2016","100%")
        );
        int rowNum = 1;

        //add for 10. Using NumberFormat for Parsing
        NumberFormat fmt = NumberFormat.getInstance(Locale.US);
        //add for  11. Setting Cell Style : Thousands Separator
        CellStyle cellStyleForThousandsSeparator = wb.createCellStyle();
        /**
         * - [apache-poi-numeric-format](https://www.baeldung.com/apache-poi-numeric-format) 另外一种写法
         * DataFormat dataFormat = wb.createDataFormat();
         * cellStyleForThousandsSeparator.setDataFormat(dataFormat.getFormat("#,##0.00"));
         */
        short builtinFormat = (short) BuiltinFormats.getBuiltinFormat("#,##0.00");
        cellStyleForThousandsSeparator.setDataFormat(builtinFormat);

        //百分比
        CellStyle cellStyleForPercentage  = wb.createCellStyle();
        cellStyleForPercentage.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0.00%"));

        for (List<String> ls : listVal) {
            Row rows = wbSheet.createRow(rowNum);
            rowNum++;
            rows.createCell(0).setCellValue(ls.get(0));
            rows.createCell(1).setCellValue(ls.get(1));

            /*10. Using NumberFormat for Parsing*/
            Number number = fmt.parse(ls.get(2));
            /*rows.createCell(2).setCellValue(number.doubleValue()); //Using NumberFormat for Parsing */

            //11. Setting Cell Style
            Cell cell = rows.createCell(2);
            cell.setCellStyle(cellStyleForThousandsSeparator);
            cell.setCellValue(number.doubleValue());

            /*9. Fixing “Number Stored As Text”*/
            rows.createCell(3).setCellValue(Double.parseDouble(ls.get(3)));

            /*12. Some More Formatting*/
            try {
                int year = Integer.parseInt(ls.get(4));
                rows.createCell(4).setCellValue(year);
            } catch (NumberFormatException exception) {
                rows.createCell(4).setCellValue(ls.get(4));
            }

            /*百分比*/
            try {
                BigDecimal divide = new BigDecimal(ls.get(5).trim().replace("%", "")).divide(BigDecimal.valueOf(100));
                Cell cell5 = rows.createCell(5);
                cell5.setCellStyle(cellStyleForPercentage);
                cell5.setCellValue(divide.doubleValue());
            } catch (NumberFormatException exception) {
                rows.createCell(5).setCellValue(ls.get(5));
            }

        }


        /*out file*/
        FileOutputStream fileOutputStream = new FileOutputStream("helloPoiExcel.xlsx");
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }
}
