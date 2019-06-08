---
title: LineFormat.Pattern property (Publisher)
keywords: vbapb10.chm3408137
f1_keywords:
- vbapb10.chm3408137
ms.prod: publisher
api_name:
- Publisher.LineFormat.Pattern
ms.assetid: ba14b1d1-9c32-a58e-d842-52fc3dc985e8
ms.date: 06/08/2019
localization_priority: Normal
---


# LineFormat.Pattern property (Publisher)

Returns or sets an **[MsoPatternType](Office.MsoPatternType.md)** constant that represents the pattern applied to the specified fill or line.


## Syntax

_expression_.**Pattern**

_expression_ A variable that represents a **[LineFormat](Publisher.LineFormat.md)** object.


## Remarks

The **Pattern** property value can be one of the **MsoPatternType** constants declared in the Microsoft Office type library.


## Example

This example sets the pattern for the specified shape if the shape currently doesn't have a fill pattern. This example assumes that at least one shape exists on the first page of the active publication.

```vb
Sub ChangeFillPattern() 
 With ActiveDocument.Pages(1).Shapes(1).Fill 
 If .Pattern < msoPattern10Percent Then 
 .Patterned Pattern:=msoPattern25Percent 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]