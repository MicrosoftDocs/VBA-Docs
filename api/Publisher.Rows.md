---
title: Rows Object (Publisher)
keywords: vbapb10.chm4980735
f1_keywords:
- vbapb10.chm4980735
ms.prod: publisher
api_name:
- Publisher.Rows
ms.assetid: 31b04a41-9005-8f51-87ab-426af0e901ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows Object (Publisher)

A collection of  **[Row](Publisher.Row.md)** objects that represent the rows in a table.
 


## Example

Use the  **[Rows](Publisher.Table.Rows.md)** property of the **[Table](Publisher.Table.md)** object to return the **Rows** collection. The following example displays the number of **[Row](Publisher.Row.md)** objects in the **Rows** collection for the first table in the active document.
 

 

```vb
Sub CountRows() 
 MsgBox ActiveDocument.Pages(2).Shapes(1).Table.Rows.Count 
End Sub
```

This example sets the fill for all even-numbered rows and clears the fill for all odd-numbered rows in the specified table. This example assumes the specified shape is a table and not another type of shape.
 

 



```vb
Sub FillCellsByRow() 
 Dim shpTable As Shape 
 Dim rowTable As Row 
 Dim celTable As Cell 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
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

Use  **Rows** (index), where index is the index number, to return a single **Row** object. The index number represents the position of the row in the **Rows** collection (counting from left to right). The following example selects the third row in the specified table.
 

 



```vb
Sub SelectRows() 
 ActiveDocument.Pages(2).Shapes(1).Table.Rows(3).Cells.Select 
End Sub
```


## Methods



|Name|
|:-----|
|[Add](Publisher.Rows.Add.md)|
|[Item](Publisher.Rows.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.Rows.Application.md)|
|[Count](Publisher.Rows.Count.md)|
|[Parent](Publisher.Rows.Parent.md)|

