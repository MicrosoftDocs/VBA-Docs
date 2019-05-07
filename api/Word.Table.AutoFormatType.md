---
title: Table.AutoFormatType property (Word)
keywords: vbawd10.chm156303466
f1_keywords:
- vbawd10.chm156303466
ms.prod: word
api_name:
- Word.Table.AutoFormatType
ms.assetid: 366dbfab-f40e-b570-d174-96f4fe07a063
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.AutoFormatType property (Word)

Returns the type of automatic formatting that's been applied to the specified table. Read-only  **Long**.


## Syntax

_expression_. `AutoFormatType`

_expression_ A variable that represents a '[Table](Word.Table.md)' object.


## Remarks

This property can be one of the  **WdTableFormat** constants. Use the **AutoFormat** method to apply automatic formatting to a table.


## Example

This example formats the first table in the active document to use the Classic 1 AutoFormat if the current format is Simple 1, Simple 2, or Simple 3.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 If ActiveDocument.Tables(1).AutoFormatType <= wdTableFormatSimple3 Then 
 ActiveDocument.Tables(1).AutoFormat _ 
 Format:=wdTableFormatClassic1 
 End If 
End If
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]