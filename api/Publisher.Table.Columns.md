---
title: Table.Columns property (Publisher)
keywords: vbapb10.chm4784131
f1_keywords:
- vbapb10.chm4784131
ms.prod: publisher
api_name:
- Publisher.Table.Columns
ms.assetid: fb55ba62-64a4-2221-3cc7-b349dc2f6934
ms.date: 06/14/2019
localization_priority: Normal
---


# Table.Columns property (Publisher)

Returns a **[Columns](Publisher.Columns.md)** collection that represents all the columns of the specified table.


## Syntax

_expression_.**Columns**

_expression_ A variable that represents a **[Table](Publisher.Table.md)** object.


## Example

This example enters a bold number into each cell in the specified table. This example assumes that the specified shape is a table and not another type of shape.

```vb
Sub CountCellsByColumn() 
 Dim shpTable As Shape 
 Dim colTable As Column 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 intCount = 1 
 
 Set shpTable = ActiveDocument.Pages(2).Shapes(1) 
 
 'Loops through each column in the table 
 For Each colTable In shpTable.Table.Columns 
 
 'Loops through each cell in the column 
 For Each celTable In colTable.Cells 
 With celTable.Text 
 .Text = intCount 
 .ParagraphFormat.Alignment = _ 
 pbParagraphAlignmentCenter 
 .Font.Bold = msoTrue 
 intCount = intCount + 1 
 End With 
 Next celTable 
 Next colTable 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]