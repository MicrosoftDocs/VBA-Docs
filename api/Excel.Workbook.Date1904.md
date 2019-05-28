---
title: Workbook.Date1904 property (Excel)
keywords: vbaxl10.chm199095
f1_keywords:
- vbaxl10.chm199095
ms.prod: excel
api_name:
- Excel.Workbook.Date1904
ms.assetid: 0556311c-4e45-aea3-e922-24a5830b19d4
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Date1904 property (Excel)

**True** if the workbook uses the 1904 date system. Read/write **Boolean**.


## Syntax

_expression_.**Date1904**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example causes Microsoft Excel to use the 1904 date system for the active workbook.

```vb
ActiveWorkbook.Date1904 = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]