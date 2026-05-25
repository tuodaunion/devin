package com.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Calendar;
import java.util.Date;

/**
 * 测试数据生成器：在指定目录下生成多个包含"销售数据"工作表的 .xlsx 文件，
 * 用于验证 ExcelMerger 的合并和排序功能。
 *
 * 使用方式：
 *   java -cp target/excel-merger-1.0.0.jar com.excel.TestDataGenerator [输出目录]
 */
public class TestDataGenerator {

    public static void main(String[] args) throws IOException {
        String outputDir = (args.length > 0) ? args[0] : System.getProperty("user.dir");
        Path dirPath = Paths.get(outputDir);

        System.out.println("正在生成测试数据到: " + dirPath.toAbsolutePath());

        // 文件1：正常数据（3行）
        createTestFile(dirPath, "销售报表_北京.xlsx", new Object[][] {
                {makeDate(2024, 3, 15), "北京店", "笔记本电脑", 5.0, 29999.50},
                {makeDate(2024, 1, 10), "北京店", "机械键盘", 20.0, 5999.00},
                {makeDate(2024, 5, 20), "北京店", "显示器", 8.0, 19992.00},
        });

        // 文件2：正常数据（3行）
        createTestFile(dirPath, "销售报表_上海.xlsx", new Object[][] {
                {makeDate(2024, 4, 5),  "上海店", "平板电脑", 12.0, 47988.00},
                {makeDate(2024, 2, 28), "上海店", "无线鼠标", 50.0, 4950.00},
                {makeDate(2024, 6, 1),  "上海店", "耳机", 30.0, 8970.00},
        });

        // 文件3：正常数据（2行）
        createTestFile(dirPath, "销售报表_广州.xlsx", new Object[][] {
                {makeDate(2024, 2, 14), "广州店", "手机壳", 100.0, 2990.00},
                {makeDate(2024, 7, 8),  "广州店", "充电器", 45.0, 4455.00},
        });

        // 文件4：没有"销售数据"工作表（应被跳过）
        createFileWithoutTargetSheet(dirPath, "其他数据.xlsx");

        System.out.println("\n测试文件生成完毕！共 4 个文件。");
        System.out.println("其中[其他数据.xlsx]不包含[销售数据]工作表，应被跳过。");
        System.out.println("预期合并后数据共 8 行，按销售日期升序排列。");
    }

    private static void createTestFile(Path dirPath, String fileName, Object[][] data) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("销售数据");

            // 创建日期格式
            CellStyle dateStyle = workbook.createCellStyle();
            CreationHelper helper = workbook.getCreationHelper();
            dateStyle.setDataFormat(helper.createDataFormat().getFormat("yyyy-MM-dd"));

            // 表头
            Row header = sheet.createRow(0);
            String[] headers = {"销售日期", "门店", "商品名称", "数量", "金额"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }

            // 数据
            for (int r = 0; r < data.length; r++) {
                Row row = sheet.createRow(r + 1);
                for (int c = 0; c < data[r].length; c++) {
                    Cell cell = row.createCell(c);
                    Object val = data[r][c];
                    if (val instanceof Date) {
                        cell.setCellValue((Date) val);
                        cell.setCellStyle(dateStyle);
                    } else if (val instanceof Double) {
                        cell.setCellValue((Double) val);
                    } else {
                        cell.setCellValue(val.toString());
                    }
                }
            }

            Path filePath = dirPath.resolve(fileName);
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
            System.out.println("  已创建: " + fileName + " (" + data.length + " 行数据)");
        }
    }

    private static void createFileWithoutTargetSheet(Path dirPath, String fileName) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("其他工作表");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("这不是销售数据");

            Path filePath = dirPath.resolve(fileName);
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
            System.out.println("  已创建: " + fileName + " (无[销售数据]工作表)");
        }
    }

    private static Date makeDate(int year, int month, int day) {
        Calendar cal = Calendar.getInstance();
        cal.set(year, month - 1, day, 0, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        return cal.getTime();
    }
}
