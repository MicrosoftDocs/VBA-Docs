---
title: CubeFields object (Excel)
keywords: vbaxl10.chm669072
f1_keywords:
- vbaxl10.chm669072
ms.prod: excel
api_name:
- Excel.CubeFields
ms.assetid: cfb7b4f4-e9c3-45a3-daa4-fe4d3c52fb1f
ms.date: 03/29/2019
localization_priority: Normal
---


# CubeFields object (Excel)

A collection of all **[CubeField](Excel.CubeField.md)** objects in a PivotTable report that is based on an OLAP cube. Each **CubeField** object represents a hierarchy or measure field from the cube.


## Example

Use the **[CubeFields](Excel.PivotTable.CubeFields.md)** property of the **PivotTable** object to return the **CubeFields** collection. The following example creates a list of cube field names of the data fields in the first OLAP-based PivotTable report on Sheet1.

```vb
Set objNewSheet = Worksheets.Add 
intRow = 1 
For Each objCubeFld In _ 
 Worksheets("Sheet1").PivotTables(1).CubeFields 
 If objCubeFld.Orientation = xlDataField Then 
 objNewSheet.Cells(intRow, 1).Value = objCubeFld.Name 
 intRow = intRow + 1 
 End If 
Next objCubeFld
```

<br/>

Use **CubeFields** (_index_), where _index_ is the cube field's index number, to return a single **CubeField** object. The following example determines the name of the second cube field in the first PivotTable report on the active worksheet.

```vb
strAlphaName = _ 
 ActiveSheet.PivotTables(1).CubeFields(2).Name
```


## Methods

- [AddSet](Excel.CubeFields.AddSet.md)
- [GetMeasure](Excel.cubefields.getmeasure.md)

## Properties

- [Application](Excel.CubeFields.Application.md)
- [Count](Excel.CubeFields.Count.md)
- [Creator](Excel.CubeFields.Creator.md)
- [Item](Excel.CubeFields.Item.md)
- [Parent](Excel.CubeFields.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]