# Excel 批量合并工具

批量读取当前目录下所有 `.xlsx` 文件中的"销售数据"工作表（A2:E100），合并后按"销售日期"升序排序，输出到 `合并后的销售数据.xlsx`。

## 环境要求

- Java 11+
- Maven 3.6+

## 快速开始

```bash
# 编译打包
mvn clean package -q

# 运行（在包含 .xlsx 文件的目录下执行）
java -jar target/excel-merger-1.0.0.jar

# 或指定目录
java -jar target/excel-merger-1.0.0.jar /path/to/excel/files
```

## 功能特性

- 自动扫描当前目录下所有 `.xlsx` 文件
- 仅读取名为"销售数据"的工作表
- 提取 A2:E100 区域数据（跳过表头行）
- 按"销售日期"列（A列）升序排序
- 完善的异常处理（文件不存在、无此工作表、数据为空、格式错误）
- 详细的中文日志输出

## 项目结构

```
├── pom.xml                                    # Maven 配置
├── README.md
└── src/main/java/com/excel/ExcelMerger.java   # 主程序
```
