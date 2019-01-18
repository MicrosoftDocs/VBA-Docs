---
title: PivotTables object (Excel)
keywords: vbaxl10.chm237072
f1_keywords:
- vbaxl10.chm237072
ms.prod: excel
api_name:
- Excel.PivotTables
ms.assetid: 5beb33ac-a0fb-3f78-8fdc-d05719512214
ms.date: 06/08/2017
localization_priority: Priority
---


# PivotTables object (Excel)

A collection of all the  **[PivotTable](Excel.PivotTable.md)** objects in the specified workbook.


## Remarks


 **Note**  The [Workbook.PivotTables](Excel.workbook.pivottables.md) property (which is new for Office) does not return all the **PivotTable** objects in the workbook; instead it returns only those associated with decoupled PivotCharts. However, [Worksheet.PivotTables](Excel.Worksheet.PivotTables.md) returns all the **PivotTable** objects in the worksheet, irrespective of whether they are associated with decoupled PivotCharts.

Because PivotTable report programming can be complex, it's generally easiest to record PivotTable report actions and then revise the recorded code.


## Example

Use the  **[PivotTables](Excel.Worksheet.PivotTables.md)** method to return the **PivotTables** collection. The following example displays the number of PivotTable reports on Sheet3.


```vb
MsgBox Worksheets("sheet3").PivotTables.Count
```

Use the  **[PivotTableWizard](Excel.Worksheet.PivotTableWizard.md)** method to create a new PivotTable report and add it to the collection. The following example creates a new PivotTable report from a Microsoft Excel database (contained in the range A1:C100).




```vb
ActiveSheet.PivotTableWizard xlDatabase, Range("A1:C100")
```

Use  **PivotTables** ( _index_ ), where _index_ is the PivotTable index number or name, to return a single **PivotTable** object. The following example makes the Year field a row field in the first PivotTable report on Sheet3.




```vb
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").Orientation = xlRowField
```


## Methods



|Name|
|:-----|
|[Add](Excel.PivotTables.Add.md)|
|[Item](Excel.PivotTables.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Excel.PivotTables.Application.md)|
|[Count](Excel.PivotTables.Count.md)|
|[Creator](Excel.PivotTables.Creator.md)|
|[Parent](Excel.PivotTables.Parent.md)|

## See also


[Excel Object Model Reference](overview/Excel/object-model.md)
