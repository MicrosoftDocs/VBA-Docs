---
title: FileExportConverter.Extensions property (Excel)
keywords: vbaxl10.chm863073
f1_keywords:
- vbaxl10.chm863073
ms.prod: excel
api_name:
- Excel.FileExportConverter.Extensions
ms.assetid: 448fdc36-4f11-1dff-98c1-797339e04ddb
ms.date: 04/26/2019
localization_priority: Normal
---


# FileExportConverter.Extensions property (Excel)

Returns the file name extensions associated with the specified **FileExportConverter** object. Read-only **String**.


## Syntax

_expression_.**Extensions**

_expression_ A variable that represents a **[FileExportConverter](Excel.FileExportConverter.md)** object.


## Example

The following example displays the file extensions for the first file converter in the **[FileExportConverters](Excel.FileExportConverters.md)** collection.


```vb
Dim fcTemp As FileExportConverter 
Set fcTemp = FileExportConverters(1) 
 
MsgBox "The file name extensions for the file converter are: " & fcTemp.Extensions
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]