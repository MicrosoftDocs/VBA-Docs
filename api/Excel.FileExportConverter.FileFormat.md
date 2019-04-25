---
title: FileExportConverter.FileFormat property (Excel)
keywords: vbaxl10.chm863075
f1_keywords:
- vbaxl10.chm863075
ms.prod: excel
api_name:
- Excel.FileExportConverter.FileFormat
ms.assetid: cdf0a922-ae9e-76b1-c8e5-228298920373
ms.date: 04/26/2019
localization_priority: Normal
---


# FileExportConverter.FileFormat property (Excel)

Returns an integer that identifies the file format associated with the specified **FileExportConverter** object. Read-only.


## Syntax

_expression_.**FileFormat**

_expression_ A variable that represents a **[FileExportConverter](Excel.FileExportConverter.md)** object.


## Example

The following example displays the file format identifier for the first file converter in the **[FileExportConverters](Excel.FileExportConverters.md)** collection.

```vb
Dim fcTemp As FileExportConverter 
Set fcTemp = FileExportConverters(1) 
 
MsgBox "The file format identifier for the file converter is: " & fcTemp.FileFormat
```

<br/>

The following example shows how to use the file format identifier as a parameter in the **[SaveAs](Excel.Workbook.SaveAs.md)** method of the **Workbook** object to save a file by using the first file converter in the **FileExportConverters** collection.

```vb
ActiveWorkbook.SaveAs _ 
 Filename:="C:\temp\myFile.xyz", _ 
 FileFormat:=Application.FileExportConverters(1).FileFormat, _ 
 CreateBackup:=False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]