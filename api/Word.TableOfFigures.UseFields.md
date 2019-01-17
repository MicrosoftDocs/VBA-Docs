---
title: TableOfFigures.UseFields property (Word)
keywords: vbawd10.chm153157641
f1_keywords:
- vbawd10.chm153157641
ms.prod: word
api_name:
- Word.TableOfFigures.UseFields
ms.assetid: 1ac7356e-fad4-1e19-1811-7df973ad74dc
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfFigures.UseFields property (Word)

 **True** if Table of Contents Entry (TC) fields are used to create a table of figures. Read/write **Boolean**.


## Syntax

 _expression_. `UseFields`

 _expression_ Required. A variable that represents a '[TableOfFigures](Word.TableOfFigures.md)' collection.


## Example

This example adds a table of figures after the selection and formats the table to compile entries with the "B" identifier.


```vb
Selection.Collapse Direction:=wdCollapseEnd 
Set myTOF = ActiveDocument.TablesOfFigures _ 
 .Add(Range:=Selection.Range) 
With myTOF 
 .UseFields = True 
 .TableId = "B" 
 .Caption = "" 
End With
```


## See also


[TableOfFigures Object](Word.TableOfFigures.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]