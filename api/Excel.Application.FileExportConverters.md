---
title: Application.FileExportConverters property (Excel)
keywords: vbaxl10.chm133318
f1_keywords:
- vbaxl10.chm133318
ms.prod: excel
api_name:
- Excel.Application.FileExportConverters
ms.assetid: 1b7289ea-344f-cc3d-ec31-04d4196533ff
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.FileExportConverters property (Excel)

Returns a **[FileExportConverters](Excel.FileExportConverters.md)** collection that represents all the file converters for saving files available to Microsoft Excel. Read-only.


## Syntax

_expression_.**FileExportConverters**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

For more information about returning a single member of a collection, see [Returning an object from a collection](../excel/Concepts/Workbooks-and-Worksheets/returning-an-object-from-a-collection-excel.md).


## Example

The following example displays the description for the first file converter in the **[FileExportConverters](Excel.FileExportConverters.md)** collection.

```vb
Dim fcTemp As FileExportConverter 
Set fcTemp = FileExportConverter(1) 
 
MsgBox fcTemp.Description
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]