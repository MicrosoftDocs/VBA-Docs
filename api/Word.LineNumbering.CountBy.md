---
title: LineNumbering.CountBy property (Word)
keywords: vbawd10.chm158466151
f1_keywords:
- vbawd10.chm158466151
ms.prod: word
api_name:
- Word.LineNumbering.CountBy
ms.assetid: 7cb90bfb-84a9-d52f-f406-7bef835744d3
ms.date: 06/08/2017
localization_priority: Normal
---


# LineNumbering.CountBy property (Word)

Returns or sets the numeric increment for line numbers. Read/write  **Long**.


## Syntax

_expression_. `CountBy`

_expression_ A variable that represents a '[LineNumbering](Word.LineNumbering.md)' object.


## Remarks

If the **CountBy** property is set to 5, every fifth line will display the line number. Line numbers are only displayed in print layout view and print preview. This property has no effect unless the **[Active](Word.LineNumbering.Active.md)** property of the **LineNumbering** object is set to **True**.


## Example

This example turns on line numbering for the active document. The line number is displayed on every fifth line and line numbering starts over for each new section.


```vb
With ActiveDocument.PageSetup.LineNumbering 
 .Active = True 
 .CountBy = 5 
 .RestartMode = wdRestartSection 
End With
```


## See also


[LineNumbering Object](Word.LineNumbering.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]