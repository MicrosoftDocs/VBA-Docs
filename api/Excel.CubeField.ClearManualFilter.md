---
title: CubeField.ClearManualFilter method (Excel)
keywords: vbaxl10.chm668098
f1_keywords:
- vbaxl10.chm668098
api_name:
- Excel.CubeField.ClearManualFilter
ms.assetid: 2dac2695-ae2c-eba9-7b22-57f21d87925a
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# CubeField.ClearManualFilter method (Excel)

The **ClearManualFilter** method provides an easy way to set the **Visible** property to **True** for all items of a PivotField in PivotTables, and to empty the **HiddenItemsList**/**VisibleItemsList** collections in OLAP PivotTables.


## Syntax

_expression_.**ClearManualFilter**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Remarks

This method is available for the **PivotField** object in PivotTables and for the **CubeField** object in the OLAP PivotTable. Calling it for a PivotField in an OLAP PivotTable will return a run-time error.

After calling this method, the **HiddenItemsList**/**HiddenItems** and **VisibleItemsList**/**VisibleItems** collections are empty.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]