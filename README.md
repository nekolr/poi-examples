# Excel for Windows 历史版本
年份 | 名称 | 版本 | BIFF 版本
---|---|---|---
1987 | Excel 2 | 2.0 | BIFF2
1990 | Excel 3 | 3.0 | BIFF3
1992 | Excel 4 | 4.0 | BIFF4
1993 | Excel 5 | 5.0 | BIFF5
1995 | Excel 95 | 7.0 | BIFF5
1997 | Excel 97 | 8.0 | BIFF8
2000 | Excel 2000 | 9.0 | BIFF8
2002 | Excel 2002 | 10.0 | BIFF8
2003 | Excel 2003 | 11.0 | BIFF8
2007 | Excel 2007 | 12.0 | 
2010 | Excel 2010 | 14.0 | 

# 关于文件后缀
格式 | 扩展名 | 说明
---|---|---
Excel 工作簿 | .xlsx | Excel 2010 和 Excel 2007 默认的基于 XML 的文件格式。 不能存储 Microsoft Visual Basic for Applications (VBA) 宏代码或 Microsoft Office Excel 4.0 宏工作表 (.xlm)。
Excel 97- Excel 2003 工作簿 | .xls | Excel 97 - Excel 2003 二进制文件格式 (BIFF8)。
Microsoft Excel 5.0/95 工作簿 | .xls | Excel 5.0/95 二进制文件格式 (BIFF5)。

> 需要注意的是：xls 格式的列不能超过 256 ，行不能超过 65536；xlsx 格式的列不能超过 16384，行不能超过 1048576。

# OLE2.0 格式
xls 后缀的文档有 Worksheet 文档和 Workbook 文档两种。Excel 4.0 及以前的版本为 Worksheet 文档，之后的版本为 Workbook 文档。其中 Worksheet 文档只包含一个 sheet，而 Workbook 文档可以包含多个 sheet，每个 Workbook 文档都包含一个全局设置（workbook globals）。

![worksheet_document](https://github.com/nekolr/poi-examples/blob/master/media/worksheet_document.png)

![workbook_document](https://github.com/nekolr/poi-examples/blob/master/media/workbook_document.png)

xls 文件使用 OLE2.0 复合文档格式进行存储。就像上面展示的，xls 文档实际上是以复合文档的形式组织在一起的，类似 apache poi 这种实现了 OLE2（POIFS）的库会以流的形式读取 Workbook 文档，Workbook 文件流会先读取 Workbook Globals Substream，然后依次读取每个 Sheet Substream。

![workbook_streams](https://github.com/nekolr/poi-examples/blob/master/media/workbook_streams.png)

## Workbook Records
Workbook 中的各种流（所有的流，包括 Substream）都由一个个 Record 组成，每个 Record 都包含特定的数据、格式等相关信息。比如 BOFRecord 记录了 Workbook 或 sheet 的开始，EOFRecord 记录了 Workbook 或 sheet 的结束。poi 库为我们抽象了各种类型的 Record，它们都在 `org.apache.poi.hssf.record` 包中。常用的有：

```
// 记录了 sheetName
BoundSheetRecord
// Workbook、Sheet 的开始
BOFRecord
// 存在单元格样式的空单元格
BlankRecord
// 布尔或错误单元格
BoolErrRecord
// 公式单元格
FormulaRecord
// 公式的计算结果单元格
StringRecord
// 文本单元格
LabelRecord
// 共用的文本单元格
LabelSSTRecord
// 数值单元格：数字单元格和日期单元格
NumberRecord
// Workbook、Sheet 的结束
EOFRecord
```

# Office Open XML 格式
Office Open XML（OOXML）是由 Microsoft 开发的一种以 XML 为基础并以 ZIP 格式压缩的电子文件规范，支持文件、表格、备忘录、幻灯片等文件格式。OOXML 在 2006 年 12 月成为了 ECMA 规范的一部分，编号为 ECMA-376，并于 2008 年 4 月通过国际标准化组织的表决，在两个月后公布为 ISO／IEC 29500 国际标准。实际上，微软公司发表的 OOXML 使用了许多非标准的规范，这造成了与其他办公软件（如 LibreOffice）发生不兼容或内容偏移的情形，微软这么做的目的是让 Microsoft Office 保持市场优势，毕竟作为 OOXML 的竞争对手 OpenDocument Format（开放文档格式，ODF）的目的就是取代私有专利文件格式，使得组织或个人不会因为文件格式而被厂商套牢。

为了验证文档是否真的使用 zip 进行了压缩，我们可以新建一个 xlsx 后缀的 excel 文档，然后使用压缩工具打开：

![zip_decompress](https://github.com/nekolr/poi-examples/blob/master/media/zip_decompress.png)

# 关于 POI
在操作 xls 文件时，poi 提供了两种编程模型，一种是 usermodel，即用户模型，另一种是 eventusermodel，即事件-用户模型。用户模型将 excel 文件抽象成了我们熟悉的诸如 Workbook、Sheet、Row、Cell 等结构，整个文档以一组对象的形式（内存树，Memory Tree）保存到内存中，使用 DOM 进行 excel 的解析，因此在文档较大时很容易发生 OOM，但是可读可写。事件-用户模型要求用户熟悉文件格式的底层结构，它的操作风格类似于 XML 的 SAX API 和 AWT 的事件模型，对于 CPU 和内存的消耗较小，但使用复杂，且无法进行写操作。

在操作 xlsx 文件时，poi 除了提供以上两种编程模型外，还提供了一种基于 XSSF 的低内存占用的 API：SXSSF，在兼容 XSSF 的同时，还能够应对大数据量和内存空间有限的情况。SXSSF 每次获取的行数是在一个数值范围内，这个范围被称为“滑动窗口”，在这个窗口内的数据均存在于内存中，超出这个窗口大小时，数据会被写入磁盘，由此控制内存使用。同时，滑动窗口的行数可以设定成自动增长的，它可以根据需要周期地调用 flushRow(int keepRows) 来进行修改。

![poi 编程模型](https://github.com/nekolr/poi-examples/blob/master/media/poi_features.png)

# 参考
> [Microsoft Excel File Format](https://www.openoffice.org/sc/excelfileformat.pdf)