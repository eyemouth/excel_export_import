package zyx.export.impl;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import zyx.export.domain.common.ExportCell;
import zyx.export.exception.FileExportException;
import zyx.export.service.IFileExportor;
import zyx.util.ReflectionUtils;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Created by Zhuwen on 2017/8/9.
 */
public class NewExcelExportor implements IFileExportor {
    @Override
    public Workbook getExportResult(List<?> data, List<ExportCell> exportCells) throws FileExportException {
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        Sheet sheet = workbook.createSheet();
        Row titleRow = sheet.createRow(0);
        createTitleRow(workbook, titleRow, exportCells, sheet);
        if (List.class.isAssignableFrom(data.getClass())) {
            if (!data.isEmpty()) {
                if (data.get(0) instanceof Map) {
                    createContentRowsByMap(workbook, (List<Map>) data, exportCells, sheet);
                } else {
                    createContentRowsByBean(workbook, (List<Object>) data, exportCells, sheet);
                }
            }
        } else {
            throw new FileExportException("传入的data数据格式有误，请检查是否属于list");
        }
        return workbook;
    }

    private static void createTitleRow(SXSSFWorkbook workbook, Row row, List<ExportCell> exportCells, Sheet sheet) {
        CellStyle style = createHeadStyle(workbook);
        row.setHeightInPoints(16.0F);
        int i = 0;
        for (ExportCell exportCell : exportCells) {
            int columnWidth = getStrLenth(exportCell.getTitle())*256;
            sheet.setColumnWidth(i, columnWidth);
            Cell cell = row.createCell(i);
            cell.setCellValue(exportCell.getTitle());
            cell.setCellStyle(style);
            cell.setCellType(HSSFCell.ENCODING_UTF_16);
            ++i;
        }
    }

    /**
     * 动态计算列宽
     *
     * @param str 列表内容字符串
     * @return
     */
    private static int getStrLenth(String str)
    {
        int len = Math.round(str.getBytes().length * 1.2f);
        len = len < 11 ? 11 : len;
        len = len > 20 ? 20 : len;
        return len;
    }

    private static CellStyle createHeadStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment((short) 2);
        style.setVerticalAlignment((short) 1);
        style.setFillPattern((short) 1);
        style.setBorderBottom((short) 1);
        style.setBorderLeft((short) 1);
        style.setBorderRight((short) 1);
        style.setBorderTop((short) 1);
        style.setWrapText(true);
        style.setFillBackgroundColor((short) 13);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.ALIGN_LEFT);
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);

        String fontName = "宋体";
        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setColor(HSSFColor.BLACK.index);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints((short) 10);//字体
        style.setFont(font);


        return style;
    }

    private Workbook createContentRowsByMap(Workbook workbook, List<Map> dataList, List<ExportCell> exportCells, Sheet sheet) throws FileExportException {
        if (!dataList.isEmpty()) {
            int rowNum = 1;
            CellStyle cellStyle = createCellStyle(workbook);
            for (Map map : dataList) {
                Row row = sheet.createRow(rowNum);
                row.setHeightInPoints(16.0F);
                for (int colNum = 0; colNum < exportCells.size(); colNum++) {
                    Cell cell = row.createCell(colNum);
                    cell.setCellStyle(cellStyle);
                    ExportCell exportCell = exportCells.get(colNum);
                    Object obj = null;
                    obj = map.get(exportCell.getAlias());
                    setCellValue(obj, cell);
                }
                rowNum++;
            }
        }
        return workbook;
    }

    private static CellStyle createCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
        style.setVerticalAlignment((short) 1);
        style.setFillForegroundColor((short) 9);
        style.setBorderBottom((short) 1);
        style.setBorderLeft((short) 1);
        style.setBorderRight((short) 1);
        style.setBorderTop((short) 1);
        Font font = workbook.createFont();
        font.setColor((short) 8);
        font.setFontHeightInPoints((short) 12);
        style.setWrapText(true);
        return style;
    }

    private static void setCellValue(Object obj, Cell cell) throws FileExportException {
        if (obj != null) {
            BigDecimal bigDecimal = null;
            String classType = obj.getClass().getName();
            if (classType.endsWith("String"))
                cell.setCellValue((String) obj);
            else if (("int".equals(classType)) || (classType.equals("java.lang.Integer")))
                cell.setCellValue(((Integer) obj).intValue());
            else if (("double".equals(classType)) || (classType.equals("java.lang.Double"))) {
                bigDecimal = new BigDecimal(((Double) obj).doubleValue());
                cell.setCellValue(bigDecimal.doubleValue());
            } else if (("float".equals(classType)) || (classType.equals("java.lang.Float"))) {
                bigDecimal = new BigDecimal(((Float) obj).floatValue());
                cell.setCellValue(bigDecimal.doubleValue());
            } else if ((classType.equals("java.util.Date")) || (classType.endsWith("Date")))
                cell.setCellValue(zyx.util.DateUtil.dateToString((Date) obj, zyx.util.DateUtil.YYYYMMDDHHMMSS));
            else if (classType.equals("java.util.Calendar"))
                cell.setCellValue((Calendar) obj);
            else if (("char".equals(classType)) || (classType.equals("java.lang.Character")))
                cell.setCellValue(obj.toString());
            else if (("long".equals(classType)) || (classType.equals("java.lang.Long")))
                cell.setCellValue(((Long) obj).longValue());
            else if (("short".equals(classType)) || (classType.equals("java.lang.Short")))
                cell.setCellValue(((Short) obj).shortValue());
            else if (classType.equals("java.math.BigDecimal")) {
                bigDecimal = (BigDecimal) obj;
                bigDecimal = new BigDecimal(bigDecimal.doubleValue());
                cell.setCellValue(bigDecimal.doubleValue());
            } else {
                throw new FileExportException("data type error !  obj is " + obj);
            }
        }
    }

    private static void createContentRowsByBean(Workbook workbook, List<Object> dataList, List<ExportCell> exportCells, Sheet sheet) throws FileExportException {
        int rowNum = 1;
        if (!dataList.isEmpty()) {
            CellStyle cellStyle = createCellStyle(workbook);
            for (Object t : dataList) {
                Row row = sheet.createRow(rowNum);
                row.setHeightInPoints(16.0F);
                for (int colNum = 0; colNum < exportCells.size(); colNum++) {
                    Cell cell = row.createCell(colNum);
                    cell.setCellStyle(cellStyle);
                    ExportCell exportCell = exportCells.get(colNum);
                    Object obj = null;
                    try {
                        obj = ReflectionUtils.executeMethod(t, ReflectionUtils.returnGetMethodName(exportCell.getAlias()));
                    } catch (Exception e) {
                        throw new FileExportException("执行executeMethod  出错 Alias is " + exportCell.getAlias() + " at " + e.getMessage());
                    }
                    setCellValue(obj, cell);
                }
                ++rowNum;
            }

        }
    }
}
