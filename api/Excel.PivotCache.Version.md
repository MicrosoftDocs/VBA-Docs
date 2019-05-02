---
title: PivotCache.Version property (Excel)
keywords: vbaxl10.chm227108
f1_keywords:
- vbaxl10.chm227108
ms.prod: excel
api_name:
- Excel.PivotCache.Version
ms.assetid: 357f61a1-7401-46c1-2a47-4172fb045cd5
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.Version property (Excel)

Returns the version of Microsoft Excel in which the PivotCache was created. Read-only **[XlPivotTableVersionList](excel.xlpivottableversionlist.md)**.


## Syntax

_expression_.**Version**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

This property returns the version of the PivotTable. Default settings and behaviors depend on the version to allow for old PivotTable object model code to run in Microsoft Office Excel 2007 with the same behaviors when the version of the PivotTable corresponds to the version of Excel for which the code was written.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]