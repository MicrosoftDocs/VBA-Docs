---
title: FillFormat.Pattern Property (Publisher)
keywords: vbapb10.chm2359558
f1_keywords:
- vbapb10.chm2359558
ms.prod: publisher
api_name:
- Publisher.FillFormat.Pattern
ms.assetid: 5b63c81e-b692-92e0-5d72-99c8d4376aff
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.Pattern Property (Publisher)

Returns an  **MsoPatternType** constant that represents the pattern applied to the specified fill or line.


## Syntax

 _expression_. **Pattern**

 _expression_ A variable that represents a  **FillFormat** object.


## Remarks

The  **Pattern** property value can be one of the ** [MsoPatternType](Office.MsoPatternType.md)** constants declared in the Microsoft Office type library.


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