package com.yqh.shop.utils;

import com.yqh.shop.model.Excel;
import org.apache.commons.collections.map.LinkedMap;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

/**
 * ExcelUtils Excel工具类
 * Created by wyh in 2018/7/5 17:30
 **/
public class ExcelUtils {
    /**
     * 解决思路：采用Apache的POI的API来操作Excel，读取内容后保存到List中，再将List转Json（推荐Linked，增删快，与Excel表顺序保持一致）
     * <p>
     * Sheet表1  ————>    List1<Map<列头，列值>>
     * Sheet表2  ————>    List2<Map<列头，列值>>
     * <p>
     * 步骤1：根据Excel版本类型创建对于的Workbook以及CellSytle
     * 步骤2：遍历每一个表中的每一行的每一列
     * 步骤3：一个sheet表就是一个Json，多表就多Json，对应一个 List
     * 一个sheet表的一行数据就是一个 Map
     * 一行中的一列，就把当前列头为key，列值为value存到该列的Map中
     *
     * @param file SSM框架下用户上传的Excel文件
     * @return Map  一个线性HashMap，以Excel的sheet表顺序，并以sheet表明作为key，sheet表转换json后的字符串作为value
     * @throws IOException
     */
//    public static Excel excel2json(MultipartFile file,String departmentId, String userPhone, String departmentName, String userId) throws IOException {
    public static Excel excel2json(MultipartFile file) throws IOException {
        System.out.println("excel2json方法执行....");

        Excel excel = new Excel();
        String[] sNames;
        // 返回的map
        LinkedHashMap<String, List> excelMap = new LinkedHashMap<>();

        // Excel列的样式，主要是为了解决Excel数字科学计数的问题
        CellStyle cellStyle;

        // Excel列的样式，主要是为了解决Excel数字科学计数的问题
        CellStyle cellStyle_D;
        // 根据Excel构成的对象
        Workbook wb;
        // 如果是2007及以上版本，则使用想要的Workbook以及CellStyle
        try {
            System.out.println(file.getOriginalFilename());
        } catch (Exception e) {
            e.printStackTrace();
        }

        if (file.getOriginalFilename().endsWith("xlsx")) {
            System.out.println("是2007及以上版本  xlsx");
            excel.setSuffix("xlsx");
            wb = new XSSFWorkbook(file.getInputStream());
            XSSFDataFormat dataFormat = (XSSFDataFormat) wb.createDataFormat();
            cellStyle = wb.createCellStyle();
            cellStyle_D = wb.createCellStyle();
            // 设置Excel列的样式为文本
            cellStyle.setDataFormat(dataFormat.getFormat("@"));
            cellStyle_D.setDataFormat(dataFormat.getFormat("yyyy/mm/dd"));
        } else if (file.getOriginalFilename().endsWith("xls")) {
            System.out.println("是2007以下版本  xls");
            excel.setSuffix("xls");
            POIFSFileSystem fs = new POIFSFileSystem(file.getInputStream());
            wb = new HSSFWorkbook(fs);
            HSSFDataFormat dataFormat = (HSSFDataFormat) wb.createDataFormat();
            cellStyle = wb.createCellStyle();
            cellStyle_D = wb.createCellStyle();
            // 设置Excel列的样式为文本
            cellStyle.setDataFormat(dataFormat.getFormat("@"));
            cellStyle_D.setDataFormat(dataFormat.getFormat("yyyy/mm/dd"));
        } else {
            return null;
        }

        // sheet表个数
        int sheetsCounts = wb.getNumberOfSheets();
        excel.setSheetCount(sheetsCounts);
        sNames = new String[sheetsCounts];
        // 遍历每一个sheet
        for (int i = 0; i < sheetsCounts; i++) {
            Sheet sheet = wb.getSheetAt(i);
            System.out.println("第" + i + "个sheet:" + sheet.toString());
            sNames[i] = sheet.getSheetName();
            // 一个sheet表对于一个List
            List list = new LinkedList();

            // 将第一行的列值作为正个json的key
            String[] cellNames;
            // 取第一行列的值作为key
            Row fisrtRow = sheet.getRow(0);
            // 如果第一行就为空，则是空sheet表，该表跳过
            if (null == fisrtRow) {
                continue;
            }
            // 得到第一行有多少列
            int curCellNum = fisrtRow.getLastCellNum();
            System.out.println("第一行的列数：" + curCellNum);
            // 根据第一行的列数来生成列头数组
            cellNames = new String[curCellNum];


            // 单独处理第一行，取出第一行的每个列值放在数组中，就得到了整张表的JSON的key
            for (int m = 0; m < curCellNum; m++) {
                Cell cell = fisrtRow.getCell(m);
                // 设置该列的样式是字符串
                cell.setCellStyle(cellStyle);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                // 取得该列的字符串值
                cellNames[m] = cell.getStringCellValue();
                //列名转key
                if ("车源类型".equals(cellNames[m])) {
                    cellNames[m] = "car_type";
                    continue;
                }
                if ("车源区域".equals(cellNames[m])) {
                    cellNames[m] = "car_region";
                    continue;
                }
                if ("供应商".equals(cellNames[m])) {
                    cellNames[m] = "company_name";
                    continue;
                }
                if ("采购员".equals(cellNames[m])) {
                    cellNames[m] = "buyer";
                    continue;
                }
                if ("订单号".equals(cellNames[m])) {
                    cellNames[m] = "purchase_number";
                    continue;
                }
                if ("省".equals(cellNames[m])) {
                    cellNames[m] = "province";
                    continue;
                }
                if ("市".equals(cellNames[m])) {
                    cellNames[m] = "city";
                    continue;
                }
                if ("区".equals(cellNames[m])) {
                    cellNames[m] = "zone";
                    continue;
                }
                if ("详细地址".equals(cellNames[m])) {
                    cellNames[m] = "address";
                    continue;
                }
                if ("备注".equals(cellNames[m])) {
                    cellNames[m] = "remark";
                    continue;
                }
                if ("品牌".equals(cellNames[m])) {
                    cellNames[m] = "car_brands";
                    continue;
                }
                if ("系列".equals(cellNames[m])) {
                    cellNames[m] = "car_series";
                    continue;
                }
                if ("型号".equals(cellNames[m])) {
                    cellNames[m] = "car_model";
                    continue;
                }
                if ("配置".equals(cellNames[m])) {
                    cellNames[m] = "car_configuration";
                    continue;
                }
                if ("外观颜色".equals(cellNames[m])) {
                    cellNames[m] = "car_exterior_color";
                    continue;
                }
                if ("内饰颜色".equals(cellNames[m])) {
                    cellNames[m] = "car_interior_color";
                    continue;
                }
                if ("车顶颜色".equals(cellNames[m])) {
                    cellNames[m] = "car_roof_color";
                    continue;
                }
                if ("车架号".equals(cellNames[m])) {
                    cellNames[m] = "car_frame_no";
                    continue;
                }
                if ("指导价".equals(cellNames[m])) {
                    cellNames[m] = "car_guide_price";
                    continue;
                }
                if ("配置价".equals(cellNames[m])) {
                    cellNames[m] = "car_configuration_price";
                    continue;
                }
                if ("批发价".equals(cellNames[m])) {
                    cellNames[m] = "car_wholesale_price";
                    continue;
                }
                if ("零售价".equals(cellNames[m])) {
                    cellNames[m] = "car_retail_price";
                    continue;
                }
                if ("配额月份".equals(cellNames[m])) {
                    cellNames[m] = "car_quota_month";
                    continue;
                }
                if ("到港日期".equals(cellNames[m])) {
                    cellNames[m] = "car_harbor_time";
                    continue;
                }
                if ("预计到店时间".equals(cellNames[m])) {
                    cellNames[m] = "car_shop_time";
                    continue;
                }
                if ("联系电话".equals(cellNames[m])) {
                    cellNames[m] = "user_phone_2";
                    continue;
                }
                excel.setCellNames(cellNames);
            }
            for (String s : cellNames) {
                System.out.print("得到第" + i + " 张sheet表的列头： " + s + ",");
            }
            System.out.println();

            // 从第二行起遍历每一行
            int rowNum = sheet.getLastRowNum();
            System.out.println("总共有 " + rowNum + " 行");
            for (int j = 1; j <= rowNum; j++) {
                // 一行数据对于一个Map
                LinkedHashMap rowMap = new LinkedHashMap();
                // 取得某一行
                Row row = sheet.getRow(j);
                int cellNum = row.getLastCellNum();
                // 遍历每一列
                for (int k = 0; k < cellNum; k++) {
                    Cell cell = row.getCell(k);


                    try {

// 保存该单元格的数据到该行中
                        Object cellValue = null;
                        switch (cell.getCellType()) {
                            case HSSFCell.CELL_TYPE_NUMERIC: // 数字

                                if (0 == cell.getCellType()) {//判断单元格的类型是否则NUMERIC类型
                                    if (HSSFDateUtil.isCellDateFormatted(cell) || isCellDateFormatted(cell)) {// 判断是否为日期类型
                                        Date date = cell.getDateCellValue();
                                        DateFormat formater = new SimpleDateFormat(
                                                "yyyy-MM-dd HH:mm");
                                        cellValue = formater.format(date);
                                    } else {
                                        cellValue = cell.getNumericCellValue() + "";
                                    }
                                }
                                break;


                            case HSSFCell.CELL_TYPE_STRING: // 字符串
                                cellValue = cell.getStringCellValue();
                                break;


                            case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                                cellValue = cell.getBooleanCellValue() + "";
                                break;


                            case HSSFCell.CELL_TYPE_FORMULA: // 公式
                                cellValue = cell.getCellFormula() + "";
                                break;


                            case HSSFCell.CELL_TYPE_BLANK: // 空值
                                cellValue = "";
                                break;


                            case HSSFCell.CELL_TYPE_ERROR: // 故障
                                cellValue = "非法字符";
                                break;


                            default:
                                cellValue = "未知类型";
                                break;
                        }

                        if ("car_region".equals(cellNames[k])) {
                            if ("东区".equals(cellValue))
                                cellValue = 1;
                            if ("西区".equals(cellValue))
                                cellValue = 2;
                            if ("南区".equals(cellValue))
                                cellValue = 3;
                            if ("北区".equals(cellValue))
                                cellValue = 4;
                        }
                        if ("car_type".equals(cellNames[k])) {
                            if ("自营".equals(cellValue))
                                cellValue = 1;
                            if ("资源".equals(cellValue))
                                cellValue = 2;
                        }
                        if ("user_phone_2".equals(cellNames[k])) {
                            if (!"".equals(cellValue)) {
                                if (!isDouble(cellValue.toString())) {
                                    cellValue = cellValue.toString();
                                } else {
                                    NumberFormat nf = NumberFormat.getInstance();
                                    nf.setGroupingUsed(false);
                                    cellValue = nf.format(Double.parseDouble(cellValue.toString()));
                                    System.err.println(nf.format(Double.parseDouble(cellValue.toString())));
                                }
                            }
                        }
                        if ("car_quota_month".equals(cellNames[k])) {
                            if (!"".equals(cellValue)) {
                                if (!isDouble(cellValue.toString())) {
                                    Date date = cell.getDateCellValue();
                                    DateFormat formater = new SimpleDateFormat(
                                            "yyyy/MM");
                                    cellValue = formater.format(date);
                                    cellValue = cellValue.toString();
                                    System.err.println(cellValue);
                                } else {
                                    NumberFormat nf = NumberFormat.getInstance();
                                    nf.setGroupingUsed(false);
                                    cellValue = nf.format(Double.parseDouble(cellValue.toString()));
                                    System.err.println(isCellDateFormatted(cell));
                                }
                            }
                        }
                        if ("purchase_number".equals(cellNames[k])) {
                            if (!"".equals(cellValue)) {
                                if (!isDouble(cellValue.toString())) {
                                    cellValue = cellValue.toString();
                                } else {
                                    NumberFormat nf = NumberFormat.getInstance();
                                    nf.setGroupingUsed(false);
                                    cellValue = nf.format(Double.parseDouble(cellValue.toString()));
                                    System.err.println(nf.format(Double.parseDouble(cellValue.toString())));
                                }
                            }
                        }
                        if ("car_harbor_time".equals(cellNames[k])) {
                            if (!"".equals(cellValue)) {
                                if (!isDouble(cellValue.toString())) {
                                    cellValue = cellValue.toString();
                                } else {
                                    NumberFormat nf = NumberFormat.getInstance();
                                    nf.setGroupingUsed(false);
                                    cellValue = nf.format(Double.parseDouble(cellValue.toString()));
                                    System.err.println(nf.format(Double.parseDouble(cellValue.toString())));
                                }
                            }
                        }
                        if ("".equals(cellValue)) {
                            cellValue = null;
                        }
                        rowMap.put(cellNames[k], cellValue);
                    } catch (Exception e) {
                        e.printStackTrace();
                        throw new RuntimeException("第" + j+ "行-" + "第" + k+1 + "列出错,请检查");
                    }

                }
                rowMap.put("resource_id", UUIDUtil.generateUUID());
//                rowMap.put("department_id",departmentId);
//                rowMap.put("user_phone",userPhone);
//                rowMap.put("department_name",departmentName);
//                rowMap.put("user_id",userId);
//                rowMap.put("create_time",new Date());
                rowMap.put("resource_no", System.currentTimeMillis() + "");
                // 保存该行的数据到该表的List中
                list.add(rowMap);
            }
            System.err.println("...........................................");
            // 将该sheet表的表名为key，List转为json后的字符串为Value进行存储
            excelMap.put(sheet.getSheetName(), list);
            excel.setExcelMap(excelMap);
        }

        System.out.println("excel2json方法结束....");

//        net.sf.json.JSONObject object = net.sf.json.JSONObject.fromObject(excelMap);
//        FileOutputStream fileOutputStream = new FileOutputStream("C:/12.txt");
//        fileOutputStream.write(object.toString().getBytes());
//        System.out.println("output finish");
        excel.setSheetNames(sNames);
        return excel;
    }


    /**
     * 判断是否为double
     *
     * @param str
     * @return
     */
    public static boolean isDouble(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException ex) {
        }
        return false;
    }

    /**
     * 判断cell类型是否为日期型
     *
     * @param Cell cell
     * @return true 是日期类型  false  否，不是日期类型
     * @throws Exception
     * @title:
     * @author xinyaoli
     * @description:
     * @date
     */
    private static boolean isCellDateFormatted(Cell cell) {
        if (cell == null) return false;
        boolean isDate = false;
        double d = cell.getNumericCellValue();
        if (DateUtil.isValidExcelDate(d)) {
            CellStyle style = cell.getCellStyle();
            if (style == null) return false;
            int i = style.getDataFormat();
            String f = style.getDataFormatString();
            isDate = DateUtil.isADateFormat(i, f);
        }
        return isDate;
    }

    public static void main(String[] args) throws Exception {
//        MultipartFile file = (MultipartFile) new File("C:/1.xlsx");

//       System.out.println( excel2json(file));
        double d = 1.234567891E10;
        System.out.println(d);
    }


}
