---
title: PivotField.EnableMultiplePageItems property (Excel)
keywords: vbaxl10.chm240151
f1_keywords:
- vbaxl10.chm240151
ms.prod: excel
api_name:
- Excel.PivotField.EnableMultiplePageItems
ms.assetid: 989fa662-cafb-00a1-effb-4a6c18327ea3
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.EnableMultiplePageItems property (Excel)

Used for specifying whether check boxes are present in the filter drop-down list for fields in the page area. Read/write **Boolean**.


## Syntax

_expression_.**EnableMultiplePageItems**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

The existing property value is retained for OLAP.

> [!NOTE] 
> In Excel 2007 or later, if you create pre-Excel 2007 OLAP PivotTables (PivotTable.Version < 3) with the **SubtotalHiddenPageItems** property of the **PivotTable** object and the **EnableMultiplePageItems** property of the **PivotField** object set to **True**, changing the state of the check boxes in the filter drop-down menu of the page area will have no effect. In this case, the filter will always be set to **All**, including the unchecked (hidden) items.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]