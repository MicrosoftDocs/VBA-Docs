---
title: Cell.HasText property (Publisher)
keywords: vbapb10.chm5111824
f1_keywords:
- vbapb10.chm5111824
ms.prod: publisher
api_name:
- Publisher.Cell.HasText
ms.assetid: b44c5d24-7ac1-a63d-6986-05ed9c91dd8e
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.HasText property (Publisher)

Returns a **Boolean** value indicating whether the specified cell contains any text. Returns **True** if the specified cell contains text. Read-only.


## Syntax

_expression_.**HasText**

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


## Example

If shape one on page one contains a table, and the first cell of the table contains text, this example displays the text in a message box.

```vb
With ActiveDocument.Pages(1).Shapes(1) 
 
 ' Check for table. 
 If .HasTable Then 
 With .Table.Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=1, EndColumn:=1).Item(1) 
 
 ' Check for text in first cell. 
 If .HasText Then 
 MsgBox "Text from first cell of table: " _ 
 & vbCr & .Text 
 Else 
 MsgBox "No text in first cell." 
 End If 
 
 End With 
 Else 
 MsgBox "No table in shape one." 
 End If 
 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]