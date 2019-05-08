---
title: PivotTable.InnerDetail property (Excel)
keywords: vbaxl10.chm235084
f1_keywords:
- vbaxl10.chm235084
ms.prod: excel
api_name:
- Excel.PivotTable.InnerDetail
ms.assetid: 385449ab-fbe2-8b69-374e-a5d374a3f76f
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.InnerDetail property (Excel)

Returns or sets the name of the field that will be shown as detail when the **ShowDetail** property is **True** for the innermost row or column field. Read/write **String**.


## Syntax

_expression_.**InnerDetail**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

This property isn't available for OLAP data sources.


## Example

This example displays the name of the field that will be shown as detail when the **ShowDetail** property is **True** for the innermost row field or column field.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox pvtTable.InnerDetail
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]