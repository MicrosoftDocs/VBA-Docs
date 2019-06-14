---
title: Table.Rows property (Publisher)
keywords: vbapb10.chm4784134
f1_keywords:
- vbapb10.chm4784134
ms.prod: publisher
api_name:
- Publisher.Table.Rows
ms.assetid: 97a543b9-a1d7-c7f8-9f3c-e08256e0b364
ms.date: 06/14/2019
localization_priority: Normal
---


# Table.Rows property (Publisher)

Returns a **[Rows](Publisher.Rows.md)** collection that represents all the table rows in a range, selection, or table.


## Syntax

_expression_.**Rows**

_expression_ A variable that represents a **[Table](Publisher.Table.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../publisher/concepts/returning-an-object-from-a-collection-publisher.md).


## Example

This example enters the fill for all even-numbered rows and clears the fill for all odd-numbered rows in the specified table. This example assumes that the specified shape is a table and not another type of shape.

```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 Set shpTable = ActiveDocument.Pages(1).Shapes _ 
 .AddTable(NumRows:=5, NumColumns:=5, Left:=100, _ 
 Top:=100, Width:=100, Height:=12) 
 For Each rowTable In shpTable.Table.Rows 
 For Each celTable In rowTable.Cells 
 If celTable.Row Mod 2 = 0 Then 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=180, Green:=180, Blue:=180) 
 Else 
 celTable.Fill.ForeColor.RGB = RGB _ 
 (Red:=255, Green:=255, Blue:=255) 
 End If 
 Next celTable 
 Next rowTable 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]