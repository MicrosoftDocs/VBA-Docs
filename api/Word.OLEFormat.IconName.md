---
title: OLEFormat.IconName property (Word)
keywords: vbawd10.chm154337287
f1_keywords:
- vbawd10.chm154337287
ms.prod: word
api_name:
- Word.OLEFormat.IconName
ms.assetid: 8894bdb0-597d-f062-e97f-1b03a7e80e26
ms.date: 06/08/2017
localization_priority: Normal
---


# OLEFormat.IconName property (Word)

Returns or sets the program file in which the icon for an OLE object is stored. Read/write  **String**.


## Syntax

_expression_. `IconName`

 _expression_ An expression that returns an '[OLEFormat](Word.OLEFormat.md)' object.


## Example

This example changes the first shape in the selection to be displayed as an icon and sets the text below the icon to the icon's file name.


```vb
Dim olefTemp As OLEFormat 
 
If Selection.ShapeRange.Count >= 1 Then 
 Set olefTemp = Selection.ShapeRange(1).OLEFormat 
 With olefTemp 
 .DisplayAsIcon = True 
 .IconLabel = .IconName 
 End With 
End If
```


## See also


[OLEFormat Object](Word.OLEFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]