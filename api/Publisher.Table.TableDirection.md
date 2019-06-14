---
title: Table.TableDirection property (Publisher)
keywords: vbapb10.chm4784135
f1_keywords:
- vbapb10.chm4784135
ms.prod: publisher
api_name:
- Publisher.Table.TableDirection
ms.assetid: ffd664a8-781f-8fdc-055c-1ea7309b3b38
ms.date: 06/14/2019
localization_priority: Normal
---


# Table.TableDirection property (Publisher)

Returns or sets a **[PbTableDirectionType](publisher.pbtabledirectiontype.md)** constant that represents whether text in a table is read from left to right or from right to left. Read/write.


## Syntax

_expression_.**TableDirection**

_expression_ A variable that represents a **[Table](Publisher.Table.md)** object.


## Return value

PbTableDirectionType


## Remarks

The **TableDirection** property value can be one of the **PbTableDirectionType** constants declared in the Microsoft Publisher type library: **pbTableDirectionLeftToRight** or **pbTableDirectionRightToLeft**.

## Example

This example enters a bold number into each cell in the specified table, and then sets the direction of the table so that the cells number from right to left. For this example to work, the specified shape must be a table.

```vb
Sub CountCellsByColumn() 
 Dim tblTable As Table 
 Dim rowTable As row 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 Set tblTable = ActiveDocument.Pages(1).Shapes(1).Table 
 
 'Loops through each row in the table 
 For Each rowTable In tblTable.Rows 
 
 'Loops through each cell in the row 
 For Each celTable In rowTable.Cells 
 With celTable.TextRange 
 intCount = intCount + 1 
 .Text = intCount 
 .ParagraphFormat.Alignment = _ 
 pbParagraphAlignmentCenter 
 .Font.Bold = msoTrue 
 End With 
 Next celTable 
 Next rowTable 
 tblTable.TableDirection = pbTableDirectionRightToLeft 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]