---
title: CubeField.CreatePivotFields method (Excel)
keywords: vbaxl10.chm668099
f1_keywords:
- vbaxl10.chm668099
ms.prod: excel
api_name:
- Excel.CubeField.CreatePivotFields
ms.assetid: 87d868d7-8836-5a0b-a4b6-1ca3165b96e0
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.CreatePivotFields method (Excel)

The **CreatePivotFields** method enables users to apply a filter to PivotFields not yet added to the PivotTable by creating the corresponding **PivotField** object.


## Syntax

_expression_.**CreatePivotFields**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Remarks

In OLAP PivotTables, PivotFields do not exist until the corresponding CubeField is added to the PivotTable. The **CreatePivotFields** method enables users to create all PivotFields of a CubeField. Users can also add filters to the PivotFields and set properties on them before the CubeField is added to the PivotTable.


## Example

```vb
Sub FilterFieldBeforeAddingItToPivotTable() 
 ActiveSheet.PivotTables("PivotTable1").CubeFields("[Date].[Fiscal]").CreatePivotFields 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Fiscal Year]").VisibleItemsList = 
 
 "[Date].[Fiscal].[Fiscal Semester]").VisibleItemsList = Array("") 
 ActiveSheet.PivotTables("PivotTable1").PivotFields( _ 
 "[Date].[Fiscal].[Fiscal Quarter]").VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Month]"). _ 
 VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Date]"). _ 
 VisibleItemsList = Array("") 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]