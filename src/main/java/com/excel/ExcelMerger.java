package com.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Excel 批量合并工具
 *
 * 功能说明：
 *   1. 扫描当前目录下所有 .xlsx 文件
 *   2. 读取每个文件中名为"销售数据"的工作表，提取 A2:E100 区域的数据
 *   3. 将所有数据合并后按"销售日期"列升序排序
 *   4. 输出到当前目录下的"合并后的销售数据.xlsx"文件
 *
 * 异常处理：
 *   - 文件不存在或无法读取
 *   - 工作表"销售数据"不存在
 *   - 数据区域为空
 *   - 单元格格式错误（日期解析失败等）
 *
 * 使用方式：
 *   mvn clean package
 *   java -jar target/excel-merger-1.0.0.jar [可选：目录路径]
 */
public class ExcelMerger {

    // ==================== 常量定义 ====================

    /** 目标工作表名称 */
    private static final String TARGET_SHEET_NAME = "销售数据";

    /** 输出文件名 */
    private static final String OUTPUT_FILE_NAME = "合并后的销售数据.xlsx";

    /** 数据起始行（跳过表头，从第2行开始，索引为1） */
    private static final int DATA_START_ROW = 1;

    /** 数据结束行（第100行，索引为99） */
    private static final int DATA_END_ROW = 99;

    /** 数据起始列（A列，索引为0） */
    private static final int DATA_START_COL = 0;

    /** 数据结束列（E列，索引为4） */
    private static final int DATA_END_COL = 4;

    /** "销售日期"列索引（默认为A列，即第0列） */
    private static final int DATE_COLUMN_INDEX = 0;

    /** 日期格式化器，用于日期比较和输出 */
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");

    // ==================== 主入口 ====================

    public static void main(String[] args) {
        System.out.println("========================================");
        System.out.println("       Excel 批量合并工具 v1.0");
        System.out.println("========================================");

        // 确定工作目录：优先使用命令行参数，否则使用当前目录
        String workDir = (args.length > 0) ? args[0] : System.getProperty("user.dir");
        Path dirPath = Paths.get(workDir);

        System.out.println("[信息] 工作目录: " + dirPath.toAbsolutePath());

        // 第一步：扫描目录下所有 .xlsx 文件（排除输出文件本身）
        List<Path> excelFiles = scanExcelFiles(dirPath);
        if (excelFiles.isEmpty()) {
            System.out.println("[警告] 当前目录下没有找到任何 .xlsx 文件，程序退出。");
            return;
        }
        System.out.println("[信息] 找到 " + excelFiles.size() + " 个 .xlsx 文件:");
        excelFiles.forEach(f -> System.out.println("       - " + f.getFileName()));

        // 第二步：逐个读取文件中的"销售数据"工作表，提取 A2:E100
        List<List<Object>> allData = new ArrayList<>();
        List<String> headerRow = null; // 保存表头（来自第一个成功读取的文件）
        int successCount = 0;
        int skipCount = 0;

        for (Path file : excelFiles) {
            System.out.println("\n[处理] 正在读取: " + file.getFileName());
            try {
                ReadResult result = readSheetData(file);
                if (result == null) {
                    // readSheetData 内部已打印原因
                    skipCount++;
                    continue;
                }
                // 保存第一次读到的表头
                if (headerRow == null && result.header != null) {
                    headerRow = result.header;
                }
                if (result.data.isEmpty()) {
                    System.out.println("  [警告] 文件中[销售数据]工作表数据区域为空，已跳过。");
                    skipCount++;
                    continue;
                }
                allData.addAll(result.data);
                successCount++;
                System.out.println("  [成功] 读取到 " + result.data.size() + " 行数据。");
            } catch (Exception e) {
                // 捕获所有异常，确保单个文件出错不影响整体流程
                System.out.println("  [错误] 读取文件失败: " + e.getMessage());
                skipCount++;
            }
        }

        System.out.println("\n========================================");
        System.out.println("[统计] 成功读取: " + successCount + " 个文件");
        System.out.println("[统计] 跳过/失败: " + skipCount + " 个文件");
        System.out.println("[统计] 合并数据总行数: " + allData.size());

        // 第三步：检查合并后的数据是否为空
        if (allData.isEmpty()) {
            System.out.println("[警告] 所有文件均无有效数据，不生成输出文件，程序退出。");
            return;
        }

        // 第四步：按"销售日期"列升序排序
        System.out.println("[信息] 正在按[销售日期]列升序排序...");
        sortByDateColumn(allData);
        System.out.println("[信息] 排序完成。");

        // 第五步：写入输出文件
        Path outputPath = dirPath.resolve(OUTPUT_FILE_NAME);
        System.out.println("[信息] 正在写入输出文件: " + outputPath.toAbsolutePath());
        writeOutputExcel(outputPath, headerRow, allData);

        System.out.println("\n========================================");
        System.out.println("[完成] 合并后的文件已保存: " + outputPath.toAbsolutePath());
        System.out.println("       共 " + allData.size() + " 行数据。");
        System.out.println("========================================");
    }

    // ==================== 扫描 Excel 文件 ====================

    /**
     * 扫描指定目录下所有 .xlsx 文件（不递归子目录）。
     * 排除输出文件本身，避免重复读取。
     *
     * @param dirPath 目录路径
     * @return .xlsx 文件路径列表，按文件名排序
     */
    private static List<Path> scanExcelFiles(Path dirPath) {
        // 校验目录是否存在且可读
        if (!Files.exists(dirPath)) {
            System.out.println("[错误] 目录不存在: " + dirPath.toAbsolutePath());
            return Collections.emptyList();
        }
        if (!Files.isDirectory(dirPath)) {
            System.out.println("[错误] 指定路径不是目录: " + dirPath.toAbsolutePath());
            return Collections.emptyList();
        }

        try (Stream<Path> stream = Files.list(dirPath)) {
            return stream
                    .filter(Files::isRegularFile)                          // 只要普通文件
                    .filter(p -> p.toString().toLowerCase().endsWith(".xlsx"))  // 只要 .xlsx
                    .filter(p -> !p.getFileName().toString().equals(OUTPUT_FILE_NAME))  // 排除输出文件
                    .filter(p -> !p.getFileName().toString().startsWith("~$"))  // 排除 Excel 临时文件
                    .sorted(Comparator.comparing(p -> p.getFileName().toString()))  // 按文件名排序
                    .collect(Collectors.toList());
        } catch (IOException e) {
            System.out.println("[错误] 扫描目录失败: " + e.getMessage());
            return Collections.emptyList();
        }
    }

    // ==================== 读取工作表数据 ====================

    /**
     * 读取单个 Excel 文件中"销售数据"工作表的 A2:E100 区域数据。
     *
     * @param filePath Excel 文件路径
     * @return ReadResult 包含表头和数据行；如果无法读取则返回 null
     */
    private static ReadResult readSheetData(Path filePath) {
        // 检查文件是否存在
        if (!Files.exists(filePath)) {
            System.out.println("  [错误] 文件不存在: " + filePath.getFileName());
            return null;
        }

        // 检查文件是否可读
        if (!Files.isReadable(filePath)) {
            System.out.println("  [错误] 文件不可读（可能被其他程序占用）: " + filePath.getFileName());
            return null;
        }

        try (FileInputStream fis = new FileInputStream(filePath.toFile());
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 查找目标工作表"销售数据"
            Sheet sheet = workbook.getSheet(TARGET_SHEET_NAME);
            if (sheet == null) {
                System.out.println("  [警告] 未找到工作表[" + TARGET_SHEET_NAME + "]，已跳过。");
                // 打印该文件中实际存在的工作表，方便排查
                System.out.print("         该文件包含的工作表: ");
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    System.out.print("[" + workbook.getSheetName(i) + "] ");
                }
                System.out.println();
                return null;
            }

            // 尝试读取表头（第1行，索引0），用于最终输出
            List<String> header = readHeaderRow(sheet);

            // 读取数据区域 A2:E100（行索引 1~99，列索引 0~4）
            List<List<Object>> data = new ArrayList<>();
            for (int rowIdx = DATA_START_ROW; rowIdx <= DATA_END_ROW; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) {
                    // 该行完全为空，跳过
                    continue;
                }

                // 读取 A~E 列（索引 0~4）的值
                List<Object> rowData = new ArrayList<>();
                boolean hasValue = false; // 标记该行是否至少有一个非空单元格

                for (int colIdx = DATA_START_COL; colIdx <= DATA_END_COL; colIdx++) {
                    Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Object value = getCellValue(cell);
                    rowData.add(value);
                    if (value != null) {
                        hasValue = true;
                    }
                }

                // 只添加至少有一个非空值的行
                if (hasValue) {
                    data.add(rowData);
                }
            }

            return new ReadResult(header, data);

        } catch (IOException e) {
            System.out.println("  [错误] 读取文件时发生 IO 异常: " + e.getMessage());
            return null;
        } catch (Exception e) {
            System.out.println("  [错误] 读取文件时发生未知异常: " + e.getMessage());
            return null;
        }
    }

    /**
     * 读取工作表的表头行（第1行）。
     *
     * @param sheet 工作表
     * @return 表头字符串列表；如果表头不存在则返回默认表头
     */
    private static List<String> readHeaderRow(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            // 没有表头行，使用默认列名
            return Arrays.asList("A", "B", "C", "D", "E");
        }

        List<String> headers = new ArrayList<>();
        for (int colIdx = DATA_START_COL; colIdx <= DATA_END_COL; colIdx++) {
            Cell cell = headerRow.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null) {
                headers.add(cell.getStringCellValue().trim());
            } else {
                headers.add("列" + (colIdx + 1));
            }
        }
        return headers;
    }

    // ==================== 单元格值提取 ====================

    /**
     * 安全地获取单元格的值，处理各种数据类型。
     * 支持：字符串、数值、日期、布尔值、公式（取缓存值）。
     *
     * @param cell 单元格对象，可能为 null
     * @return 单元格的值（String / Double / Date / Boolean），null 表示空单元格
     */
    private static Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        try {
            switch (cell.getCellType()) {
                case STRING:
                    // 字符串类型：去除首尾空格
                    String strValue = cell.getStringCellValue();
                    return (strValue != null && !strValue.trim().isEmpty()) ? strValue.trim() : null;

                case NUMERIC:
                    // 数值类型：需要区分日期和普通数字
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 日期类型
                        return cell.getDateCellValue();
                    } else {
                        // 普通数值
                        return cell.getNumericCellValue();
                    }

                case BOOLEAN:
                    return cell.getBooleanCellValue();

                case FORMULA:
                    // 公式类型：尝试获取缓存的计算结果
                    return getFormulaValue(cell);

                case BLANK:
                    return null;

                case ERROR:
                    System.out.println("    [警告] 单元格包含错误值，已跳过。行: "
                            + (cell.getRowIndex() + 1) + ", 列: " + (char) ('A' + cell.getColumnIndex()));
                    return null;

                default:
                    return null;
            }
        } catch (Exception e) {
            // 格式错误等异常：记录警告并返回 null，不中断流程
            System.out.println("    [警告] 读取单元格值失败 (行: " + (cell.getRowIndex() + 1)
                    + ", 列: " + (char) ('A' + cell.getColumnIndex()) + "): " + e.getMessage());
            return null;
        }
    }

    /**
     * 获取公式单元格的缓存值。
     *
     * @param cell 公式单元格
     * @return 缓存的计算结果
     */
    private static Object getFormulaValue(Cell cell) {
        try {
            switch (cell.getCachedFormulaResultType()) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    }
                    return cell.getNumericCellValue();
                case STRING:
                    return cell.getStringCellValue();
                case BOOLEAN:
                    return cell.getBooleanCellValue();
                default:
                    return null;
            }
        } catch (Exception e) {
            // 公式结果获取失败，返回公式文本作为备用
            return cell.getCellFormula();
        }
    }

    // ==================== 数据排序 ====================

    /**
     * 按"销售日期"列（第一列）升序排序。
     * 排序逻辑：
     *   - 日期类型优先，按时间升序
     *   - 字符串类型按字典序
     *   - null 值排到最后
     *
     * @param data 待排序的数据列表
     */
    private static void sortByDateColumn(List<List<Object>> data) {
        data.sort((row1, row2) -> {
            Object val1 = row1.get(DATE_COLUMN_INDEX);
            Object val2 = row2.get(DATE_COLUMN_INDEX);

            // null 值排到最后
            if (val1 == null && val2 == null) return 0;
            if (val1 == null) return 1;
            if (val2 == null) return -1;

            // 两者都是 Date 类型
            if (val1 instanceof Date && val2 instanceof Date) {
                return ((Date) val1).compareTo((Date) val2);
            }

            // 两者都是 String 类型（可能是日期字符串）
            if (val1 instanceof String && val2 instanceof String) {
                // 尝试解析为日期进行比较
                Date date1 = tryParseDate((String) val1);
                Date date2 = tryParseDate((String) val2);
                if (date1 != null && date2 != null) {
                    return date1.compareTo(date2);
                }
                // 无法解析为日期，按字符串排序
                return ((String) val1).compareTo((String) val2);
            }

            // 混合类型：Date 排在 String 前面
            if (val1 instanceof Date) return -1;
            if (val2 instanceof Date) return 1;

            // 其他类型：转为字符串比较
            return val1.toString().compareTo(val2.toString());
        });
    }

    /**
     * 尝试将字符串解析为日期（支持 yyyy-MM-dd 和 yyyy/MM/dd 格式）。
     *
     * @param dateStr 日期字符串
     * @return 解析后的 Date 对象；解析失败返回 null
     */
    private static Date tryParseDate(String dateStr) {
        String[] formats = {"yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", "yyyyMMdd"};
        for (String fmt : formats) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(fmt);
                sdf.setLenient(false); // 严格模式，防止无效日期
                return sdf.parse(dateStr);
            } catch (Exception e) {
                // 尝试下一个格式
            }
        }
        return null;
    }

    // ==================== 写入输出文件 ====================

    /**
     * 将合并后的数据写入新的 Excel 文件。
     *
     * @param outputPath 输出文件路径
     * @param headers    表头列表（可能为 null）
     * @param data       数据行列表
     */
    private static void writeOutputExcel(Path outputPath, List<String> headers, List<List<Object>> data) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(TARGET_SHEET_NAME);

            // 创建表头样式：加粗、居中
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);

            // 创建日期格式样式
            CellStyle dateStyle = workbook.createCellStyle();
            CreationHelper createHelper = workbook.getCreationHelper();
            dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd"));

            // 写入表头（第1行）
            int currentRow = 0;
            if (headers != null && !headers.isEmpty()) {
                Row headerRowObj = sheet.createRow(currentRow++);
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRowObj.createCell(i);
                    cell.setCellValue(headers.get(i));
                    cell.setCellStyle(headerStyle);
                }
            }

            // 写入数据行
            for (List<Object> rowData : data) {
                Row row = sheet.createRow(currentRow++);
                for (int colIdx = 0; colIdx < rowData.size(); colIdx++) {
                    Cell cell = row.createCell(colIdx);
                    Object value = rowData.get(colIdx);

                    if (value == null) {
                        cell.setBlank();
                    } else if (value instanceof Date) {
                        cell.setCellValue((Date) value);
                        cell.setCellStyle(dateStyle);
                    } else if (value instanceof Double) {
                        cell.setCellValue((Double) value);
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    } else {
                        cell.setCellValue(value.toString());
                    }
                }
            }

            // 自动调整列宽
            for (int colIdx = DATA_START_COL; colIdx <= DATA_END_COL; colIdx++) {
                sheet.autoSizeColumn(colIdx);
                // 设置最小宽度为 12 个字符（避免列太窄）
                if (sheet.getColumnWidth(colIdx) < 12 * 256) {
                    sheet.setColumnWidth(colIdx, 12 * 256);
                }
            }

            // 写入文件
            try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                workbook.write(fos);
            }

            System.out.println("[信息] 文件写入成功。");

        } catch (IOException e) {
            System.out.println("[错误] 写入输出文件失败: " + e.getMessage());
        } catch (Exception e) {
            System.out.println("[错误] 生成输出文件时发生未知异常: " + e.getMessage());
        }
    }

    // ==================== 内部辅助类 ====================

    /**
     * 读取结果封装类，包含表头和数据。
     */
    private static class ReadResult {
        /** 表头行 */
        final List<String> header;
        /** 数据行 */
        final List<List<Object>> data;

        ReadResult(List<String> header, List<List<Object>> data) {
            this.header = header;
            this.data = data;
        }
    }
}
