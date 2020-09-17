# CsvToExcel

csvをExcelに貼り付けて表をいい感じにするやつ

## いる

.NET Framework 4.7.2

## 使い方

```
set pg="exeのパス"
set csvDir="csvファイルのあるディレクトリ"
set excelDir="excelファイルを出力するディレクトリ"
set nameBase="excelファイルのベース部分（※後ろにタイムスタンプと拡張子が付きます）"
%pg% %csvDir% %excelDir% %nameBase%
```

## 使った

- ClosedXML https://www.nuget.org/packages/ClosedXML/
- DocumentFormat.OpenXml https://www.nuget.org/packages/DocumentFormat.OpenXml/
- ExcelNumberFormat https://www.nuget.org/packages/ExcelNumberFormat/
