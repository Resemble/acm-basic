package class_06;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.*;

import java.io.File;
import java.io.IOException;
import java.util.Locale;
import java.util.regex.Pattern;

/**
 * @author ranbo
 * @version V1.0
 * @Title:
 * @Package convert
 * @Description:
 * @date 2019-10-27 16:57
 */
public class ExcelConvert {

    static Sheet sheet1;
    static Sheet sheet2;
    static WritableSheet excelSheet;
    static WritableSheet excelSheet1;
    static WritableSheet excelSheet2;

    public static void main(String[] args) throws IOException, BiffException, WriteException {

        File inputStream1 = new File("/Users/ranbo/Downloads/股票代码.xls");
        File inputStream2 = new File("/Users/ranbo/Downloads/重大重组事件.xls");
        Workbook workbook1 = Workbook.getWorkbook(inputStream1);
        sheet1 = workbook1.getSheet(0);
        sheet2 = Workbook.getWorkbook(inputStream2).getSheet(0);
        System.out.println("sheet1.getColumns():" + sheet1.getColumns());
        System.out.println("sheet1.getRows():" + sheet1.getRows());
        System.out.println("sheet2.getColumns():" + sheet2.getColumns());
        System.out.println("sheet2.getRows():" + sheet2.getRows());
        /** 每个值去匹配，如果找到写入新的 excel */
        //        for (int i = 1; i < sheet1.getRows(); i++) {
        //            for (int j = 0; j < sheet1.getColumns(); j++) {
        //                Cell cell = sheet1.getCell(j, i);
        //                System.out.println(cell.getContents());
        //            }
        //        }
        //        for (int i = 1; i < sheet2.getRows(); i++) {
        //            for (int j = 0; j < sheet2.getColumns(); j++) {
        //                Cell cell = sheet2.getCell(j, i);
        //                System.out.println(cell.getContents());
        //            }
        //        }

        // j:10
        //交易标的

        //        for (int j = 0; j < sheet2.getColumns(); j++) {
        //            System.out.println("j:" + j);
        //            Cell cell = sheet2.getCell(j, 0);
        //            System.out.println(cell.getContents());
        //        }


        File file = new File("/Users/ranbo/Downloads/合并后.xls");
        if (file.exists()) {
            file.delete();
            file = new File("/Users/ranbo/Downloads/合并后.xls");
        }
        WorkbookSettings wbSettings = new WorkbookSettings();

        wbSettings.setLocale(new Locale("en", "EN"));

        WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("匹配上", 0);
        excelSheet = workbook.getSheet(0);

        workbook.createSheet("没有匹配上的股票代码", 1);
        excelSheet1 = workbook.getSheet(1);

        workbook.createSheet("没有匹配上的重大重组事件", 2);
        excelSheet2 = workbook.getSheet(2);


        int[] excel1NotFoundRow = new int[sheet1.getRows()];
        int[] excel2NotFoundRow = new int[sheet2.getRows()];

        /** createLabel */
        write(0, 0, 0);

        /** createContent */
        int cnt = 0;
        for (int j = 1; j < sheet1.getRows(); j++) {
            Cell cell = sheet1.getCell(1, j);
            //            System.out.println(cell.getContents().trim());
            String content1 = cell.getContents().trim().replace("*", "");
            for (int i = 1; i < sheet2.getRows(); i++) {
                Cell cell2 = sheet2.getCell(10, i);
                //                System.out.println(cell2.getContents().trim());

                boolean isMatch = Pattern.matches(".*" + content1 + ".*", cell2.getContents());
                if (isMatch) {
                    cnt++;
                    System.out.println("第" + cnt + "个 j:" + j + " i:" + i);
                    System.out.println(
                        cell.getContents().trim() + " match " + cell2.getContents().trim());

                    excel1NotFoundRow[j] = 1;
                    excel2NotFoundRow[i] = 1;
                    write(cnt, j, i);

                }
            }
        }



        //        createLabel(excelSheet);
        //        createContent(excelSheet);


        int writeRow = 0;
//
        for (int j = 0; j < sheet2.getRows(); j++) {
            if (excel2NotFoundRow[j] == 0) {
                for (int i = 0; i < sheet2.getColumns(); i++) {
                    Cell cell2 = sheet2.getCell(i, j);
                    Label label = new Label(i, writeRow, cell2.getContents());
                    excelSheet2.addCell(label);
                }
                writeRow++;
            }
        }

        writeRow = 0;
        for (int j = 0; j < sheet1.getRows(); j++) {
            if (excel1NotFoundRow[j] == 0) {
                for (int i = 0; i < sheet1.getColumns(); i++) {
                    Cell cell1 = sheet1.getCell(i, j);
                    Label label = new Label(i, writeRow, cell1.getContents());
                    excelSheet1.addCell(label);
                }
                writeRow++;
            }
        }




        workbook.write();

        workbook1.close();
        workbook.close();
    }

    public static void write(int writeRow, int sheet1Row, int sheet2Row) throws WriteException {
        int column = 0;

        for (int i = 0; i < sheet2.getColumns(); i++) {
            Cell cell2 = sheet2.getCell(i, sheet2Row);
            Label label = new Label(column++, writeRow, cell2.getContents());
            excelSheet.addCell(label);
            if (i == 10) {
                for (int j = 0; j < sheet1.getColumns(); j++) {
                    Cell cell1 = sheet1.getCell(j, sheet1Row);
                    Label label2 = new Label(column++, writeRow, cell1.getContents());
                    excelSheet.addCell(label2);
                }
            }
        }
    }



}
