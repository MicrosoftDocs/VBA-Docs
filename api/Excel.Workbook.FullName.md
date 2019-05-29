---
title: Workbook.FullName property (Excel)
keywords: vbaxl10.chm199102
f1_keywords:
- vbaxl10.chm199102
ms.prod: excel
api_name:
- Excel.Workbook.FullName
ms.assetid: 83f45d15-b009-f304-ca53-4daa80c06562
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.FullName property (Excel)

Returns the name of the object, including its path on disk, as a string. Read-only **String**.


## Syntax

_expression_.**FullName**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example displays the path and file name of the active workbook (assuming that the workbook has been saved).

```vb
MsgBox ActiveWorkbook.FullName
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
