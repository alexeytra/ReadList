package classes;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.util.List;

public class WriteExcel {
    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillBackgroundColor(HSSFColor.LAVENDER.index);
        return style;
    }

    private static HSSFCellStyle createStyleForCell(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 11);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }

    public static void WriteToExcel(List<ListOne> l, int choose) throws IOException {

        File file = new File("List.xls");
        //file.getParentFile().mkdirs();
        HSSFWorkbook workbook = null;

        if (file.exists()){
            //Запись если файл существует
            FileInputStream inputStream = new FileInputStream(file);
            workbook = new HSSFWorkbook(inputStream);
            write(workbook, choose, l, file);
            //
            inputStream.close();
        } else {
            workbook = new HSSFWorkbook();
            write(workbook, choose, l, file);
            //
        }

    }

    public static void write(HSSFWorkbook workbook, int choose, List<ListOne> l, File file){
        FileOutputStream outFile = null;
        HSSFSheet sheet;
        switch (choose) {
            case 1: sheet = workbook.createSheet("СТ отрасли"); break;
            case 2: sheet = workbook.createSheet("СТ виды работ"); break;
            case 3: sheet = workbook.createSheet("СТ подпункты"); break;
            case 4: sheet = workbook.createSheet("СТ позиции"); break;
            case 5: sheet = workbook.createSheet("СТ пункты"); break;
            default:  sheet = workbook.createSheet("error");
        }
        int rownum = 0;
        Cell cell;
        Row row;
        HSSFCellStyle style = createStyleForTitle(workbook);
        HSSFCellStyle styleforSell = createStyleForCell(workbook);

        row = sheet.createRow(rownum);

        //Название столбцов
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("discription (Описание)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Сode (Кодировка)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("value (Значение)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue("type (Тип)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(4, CellType.STRING);
        cell.setCellValue("begin (Срок с)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(5, CellType.STRING);
        cell.setCellValue("end (Срок по)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(6, CellType.STRING);
        cell.setCellValue("beginPer (Срок периода с)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(7, CellType.STRING);
        cell.setCellValue("endPer (Срок периода по)");
        cell.setCellStyle(style);
        //
        cell = row.createCell(8, CellType.STRING);
        cell.setCellValue("НПА");
        cell.setCellStyle(style);

        for (ListOne p : l) {
            rownum++;
            row = sheet.createRow(rownum);
            //
            cell = row.createCell(0, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getDescription());
            //
            cell = row.createCell(1, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getCode());
            //
            cell = row.createCell(2, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getValue());
            //
            cell = row.createCell(3, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getType());
            //
            cell = row.createCell(4, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getBegin());
            //
            cell = row.createCell(5, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getEnd());
            //
            cell = row.createCell(6, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getBeginPer());
            //
            cell = row.createCell(7, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getEndPer());
            //
            cell = row.createCell(8, CellType.STRING);
            cell.setCellStyle(styleforSell);
            cell.setCellValue(p.getNpa());
        }
        for (int i = 1; i <= 8; i++) {
            sheet.autoSizeColumn(i);
        }

        //Запись результата
        try {
            outFile = new FileOutputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(outFile);
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Все ok " + file.getAbsolutePath());
    }


}
