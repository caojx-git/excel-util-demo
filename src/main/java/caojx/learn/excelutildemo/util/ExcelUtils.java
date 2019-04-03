package caojx.learn.excelutildemo.util;

import caojx.learn.excelutildemo.annotation.ExcelCell;
import caojx.learn.excelutildemo.annotation.ExcelSheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.joda.time.DateTime;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel 工具类
 *
 * @author caojx
 * @version $Id: ExcelUtils.java,v 1.0 2019-04-03 17:12 caojx
 * @date 2019-04-03 17:12
 */
@Slf4j
public class ExcelUtils {

    /**
     * 导出Excel
     *
     * @param request
     * @param response
     * @param fileName
     * @param data
     * @param <T>
     * @throws Exception
     */
    public static <T> void exportToExcel(HttpServletRequest request, HttpServletResponse response, String fileName, List<T> data) throws Exception {
        Workbook wb = getWorkbook(data);
        String agent = request.getHeader("USER-AGENT").toLowerCase();
        try (OutputStream out = response.getOutputStream()) {
            response.setContentType("application/vnd.ms-excel");
            //浏览器兼容
            if (agent.contains("mozilla")) {
                response.setCharacterEncoding("utf-8");
                response.addHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes("gbk"), "iso8859-1"));
            } else {
                response.setHeader("content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            }
            wb.write(out);
        } catch (Exception e) {
            log.error("exportToExcel, exception:{}", e);
            throw e;
        }
    }

    /**
     * 获取工作簿
     *
     * @param list
     * @param <T>
     * @return
     */
    private static <T> Workbook getWorkbook(List<T> list) {
        Workbook wb = new SXSSFWorkbook();
        if (list != null && !list.isEmpty()) {
            try {
                makeSheet(wb, list);
            } catch (Exception e) {
                log.error(e.getMessage(), e);
            }
            return wb;
        } else {
            wb.createSheet("Sheet");
            return wb;
        }
    }

    /**
     * 创建sheet
     *
     * @param wb
     * @param list
     * @param <T>
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     */
    private static <T> void makeSheet(Workbook wb, List<T> list) throws IllegalArgumentException, IllegalAccessException {
        Object titleRow = list.get(0);
        Sheet sheet = createSheet(wb, titleRow);
        CellStyle titleStyle = createTitleStyle(wb);
        makeSheetTitle(wb, sheet, titleRow, titleStyle);
        makeSheetData(wb, sheet, list);
    }

    private static Sheet createSheet(Workbook wb, Object titleRow) {
        Class<?> baseClazz = titleRow.getClass();
        ExcelSheet excelSheet = baseClazz.getAnnotation(ExcelSheet.class);
        return excelSheet != null ? wb.createSheet(excelSheet.value()) : wb.createSheet();
    }

    /**
     * 设置样式
     *
     * @param wb
     * @return
     */
    private static CellStyle createTitleStyle(Workbook wb) {
        CellStyle titleStyle = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 13);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontName("黑体");
        titleStyle.setFont(font);
        titleStyle.setBorderTop((short) 6);
        titleStyle.setBorderBottom((short) 1);
        titleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        titleStyle.setWrapText(true);
        titleStyle.setBorderLeft(CellStyle.BORDER_THIN);
        titleStyle.setBorderRight(CellStyle.BORDER_THIN);
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        return titleStyle;
    }

    /**
     * 设置注解为@ExcelCell的值作为表头内容
     *
     * @param wb
     * @param sheet
     * @param titleRow
     * @param titleStyle
     */
    private static void makeSheetTitle(Workbook wb, Sheet sheet, Object titleRow, CellStyle titleStyle) {
        Class<?> baseClazz = titleRow.getClass();
        Row headRow = sheet.createRow(0);
        Field[] headFields = baseClazz.getDeclaredFields();
        int cHeadIndex = 0;
        Field[] var8 = headFields;
        int var9 = headFields.length;

        for (int var10 = 0; var10 < var9; ++var10) {
            Field field = var8[var10];
            ExcelCell excelCell = (ExcelCell) field.getAnnotation(ExcelCell.class);
            if (excelCell != null) {
                Cell cell = headRow.createCell(cHeadIndex);
                cell.setCellType(1);
                cell.setCellStyle(titleStyle);
                if (!"".equals(excelCell.value())) {
                    cell.setCellValue(excelCell.value());
                } else {
                    cell.setCellValue(field.getName());
                }

                sheet.setColumnWidth(cHeadIndex, (short) (35.7 * 150));
                ++cHeadIndex;
            }
        }

    }

    /**
     * 设置数据
     *
     * @param wb
     * @param sheet
     * @param list
     * @param <T>
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     */
    private static <T> void makeSheetData(Workbook wb, Sheet sheet, List<T> list) throws IllegalArgumentException, IllegalAccessException {
        Map<String, CellStyle> cellStyleMap = getCellStyleMap(wb);
        for (int rIndex = 0; rIndex < list.size(); ++rIndex) {
            Row row = sheet.createRow(1 + rIndex);
            Object rowData = list.get(rIndex);
            Class<?> clazz = rowData.getClass();
            Field[] fields = clazz.getDeclaredFields();
            int cIndex = 0;
            Field[] var10 = fields;
            int var11 = fields.length;

            for (int var12 = 0; var12 < var11; ++var12) {
                Field field = var10[var12];
                ExcelCell excelCell = (ExcelCell) field.getAnnotation(ExcelCell.class);
                if (excelCell != null) {
                    field.setAccessible(true);
                    Object value = field.get(rowData);
                    if (value != null) {
                        Cell cell = row.createCell(cIndex);
                        setCellValue(cell, value, field.getType(), excelCell, cellStyleMap);
                    }
                    ++cIndex;
                }
            }
        }

    }

    private static Map<String, CellStyle> getCellStyleMap(Workbook wb) {
        Map<String, CellStyle> cellStyleMap = new HashMap();
        CellStyle cellStyle4Num = wb.createCellStyle();
        cellStyle4Num.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
        cellStyleMap.put("NUMBER", cellStyle4Num);
        return cellStyleMap;
    }

    /**
     * 设置单元格值
     *
     * @param cell
     * @param value
     * @param fieldType
     * @param excelCell
     * @param cellStyleMap
     */
    private static void setCellValue(Cell cell, Object value, Class<?> fieldType, ExcelCell excelCell, Map<String, CellStyle> cellStyleMap) {
        String fieldTypeName = fieldType.getSimpleName();
        byte var7 = -1;
        switch (fieldTypeName.hashCode()) {
            case -1808118735:
                if (fieldTypeName.equals("String")) {
                    var7 = 0;
                }
                break;
            case -1325958191:
                if (fieldTypeName.equals("double")) {
                    var7 = 5;
                }
                break;
            case -672261858:
                if (fieldTypeName.equals("Integer")) {
                    var7 = 2;
                }
                break;
            case 104431:
                if (fieldTypeName.equals("int")) {
                    var7 = 1;
                }
                break;
            case 2086184:
                if (fieldTypeName.equals("Byte")) {
                    var7 = 14;
                }
                break;
            case 2122702:
                if (fieldTypeName.equals("Date")) {
                    var7 = 15;
                }
                break;
            case 2374300:
                if (fieldTypeName.equals("Long")) {
                    var7 = 4;
                }
                break;
            case 3039496:
                if (fieldTypeName.equals("byte")) {
                    var7 = 13;
                }
                break;
            case 3327612:
                if (fieldTypeName.equals("long")) {
                    var7 = 3;
                }
                break;
            case 64711720:
                if (fieldTypeName.equals("boolean")) {
                    var7 = 11;
                }
                break;
            case 67973692:
                if (fieldTypeName.equals("Float")) {
                    var7 = 8;
                }
                break;
            case 79860828:
                if (fieldTypeName.equals("Short")) {
                    var7 = 10;
                }
                break;
            case 97526364:
                if (fieldTypeName.equals("float")) {
                    var7 = 7;
                }
                break;
            case 109413500:
                if (fieldTypeName.equals("short")) {
                    var7 = 9;
                }
                break;
            case 1729365000:
                if (fieldTypeName.equals("Boolean")) {
                    var7 = 12;
                }
                break;
            case 1857393595:
                if (fieldTypeName.equals("DateTime")) {
                    var7 = 17;
                }
                break;
            case 2052876273:
                if (fieldTypeName.equals("Double")) {
                    var7 = 6;
                }
                break;
            case 2059094262:
                if (fieldTypeName.equals("Timestamp")) {
                    var7 = 16;
                }
                break;
            case 1438607953:
                if (fieldTypeName.equals("BigDecimal")) {
                    var7 = 18;
                }
        }

        switch (var7) {
            case 0:
                cell.setCellValue((String) value);
                break;
            case 1:
            case 2:
                cell.setCellValue((double) ((Integer) value).intValue());
                break;
            case 3:
            case 4:
                cell.setCellStyle((CellStyle) cellStyleMap.get("NUMBER"));
                cell.setCellValue((double) ((Long) value).longValue());
                break;
            case 5:
            case 6:
                cell.setCellValue(((Double) value).doubleValue());
                break;
            case 7:
            case 8:
                cell.setCellValue((double) ((Float) value).floatValue());
                break;
            case 9:
            case 10:
                cell.setCellValue((double) ((Short) value).shortValue());
                break;
            case 11:
            case 12:
                cell.setCellValue(((Boolean) value).booleanValue());
                break;
            case 13:
            case 14:
                cell.setCellValue((double) ((Byte) value).byteValue());
                break;
            case 15:
            case 16:
            case 17:
                cell.setCellValue((new DateTime(value)).toString(excelCell.timeFormat()));
                break;
            case 18:
                cell.setCellValue(((BigDecimal) value).doubleValue());
        }
    }


    /**
     * 上传Excel
     * 解析后的到一个map
     *  key作为sheet行编号
     *  value为list，list按顺序存储第0行的0list到第n行的list
     *
     * @param file 上传的文件
     * @param path 文件保存路径
     * @return
     */
    public static Map<Integer, List<List<? super Object>>> uploadExcel(MultipartFile file, String path) {
        //创建读取excel的类(区分excel2003和2007文件)
        Workbook workbook = createWorkBook(file, path);
        //excel每一行看做一个List,作为value;sheet页数作为key
        Map<Integer, List<List<? super Object>>> map = new HashMap<>();
        // 得到sheet页
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            // 得到Excel的行数
            int totalRows = sheet.getPhysicalNumberOfRows();
            // 得到Excel的列数(前提是有行数)
            int totalCells = 0;
            if (totalRows > 1 && sheet.getRow(0) != null) {
                totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
            }
            // 循环Excel行数
            List<List<? super Object>> list = new ArrayList<>();
            for (int r = 1; r < totalRows; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                // 循环Excel的列
                if (row != null) {
                    List<? super Object> valueList = new ArrayList<>();
                    for (int c = 0; c < totalCells; c++) {
                        Cell cell = row.getCell(c);
                        DecimalFormat df = new DecimalFormat("0");
                        try {
                            if (null != cell) {
                                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                    String value = df.format(cell.getNumericCellValue());
                                    valueList.add(value);
                                } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                                    if (!StringUtils.isEmpty(cell.getStringCellValue())) {
                                        valueList.add(cell.getStringCellValue());
                                    }
                                } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                    valueList.add("");
                                } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                                    valueList.add(cell.getBooleanCellValue());
                                } else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
                                    valueList.add("非法字符");
                                } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                                    valueList.add(cell.getCellFormula());
                                }
                                ;
                            }
                        } catch (Exception e) {
                            log.error("\n==========解析Excel单元格异常==========", e);
                        }
                    }
                    list.add(valueList);
                }
            }
            map.put(i, list);
        }
        return map;
    }

    /**
     * 创建工作簿
     *
     * @param file
     * @param path
     * @return
     */
    private static Workbook createWorkBook(MultipartFile file, String path) {
        multipartToFile(file, path);
        File f = createNewFile(file, path);
        Workbook workbook = null;
        try {
            InputStream is = new FileInputStream(f);
            workbook = WorkbookFactory.create(is);
            is.close();
        } catch (InvalidFormatException | IOException e) {
            log.error("\n==========文件流转换为Workbook对象异常============", e);
        }
        return workbook;
    }

    /**
     * 保存文件到指定路径
     *
     * @param multfile
     * @param path
     */
    private static void multipartToFile(MultipartFile multfile, String path) {
        File file = createNewFile(multfile, path);
        try {
            multfile.transferTo(file);
        } catch (IOException e) {
            log.error("\n上传的文件保存失败");
        }
    }

    /**
     * 获取上传的文件
     *
     * @param multfile
     * @param path
     * @return
     */
    private static File createNewFile(MultipartFile multfile, String path) {
        String fileName = multfile.getOriginalFilename();
        File file = new File(path + fileName);
        File parentFile = file.getParentFile();
        if (!parentFile.exists()) {
            parentFile.mkdirs();
        }
        try {
            file.createNewFile();
        } catch (IOException e) {
            log.error("\n新文件创建失败");
        }
        return file;
    }
}