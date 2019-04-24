---
title: CubeField.HierarchizeDistinct property (Excel)
keywords: vbaxl10.chm668104
f1_keywords:
- vbaxl10.chm668104
ms.prod: excel
api_name:
- Excel.CubeField.HierarchizeDistinct
ms.assetid: 714f85b7-2adb-0ec1-5203-ca797b21e0a8
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.HierarchizeDistinct property (Excel)

Returns or sets whether to order and remove duplicates when displaying the specified named set in a PivotTable report based on an OLAP cube. Read/write.


## Syntax

_expression_.**HierarchizeDistinct**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Return value

**Boolean**


## Remarks

**True** if the named set is displayed as ordered with duplicates removed; otherwise, **False**.

The value of this property corresponds to the setting of the **Automatically order and remove duplicates from the set** check box on the **Layout & Print** tab of the **Field Settings** dialog box for a named set in a PivotTable report based on an OLAP cube.

This property returns an error if the **[CubeFieldType](Excel.CubeField.CubeFieldType.md)** property of the specified **CubeField** object is not **xlSet** (**[XlCubeFieldType](excel.xlcubefieldtype.md)** enumeration).


## Example

The following code example sets the **HierarchizeDistinct** property to **True** to order and remove duplicates from the specified named set.

```vb
ActiveSheet.PivotTables("PivotTable1").CubeFields("[Summary P&L]"). _ 
 HierarchizeDistinct = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]