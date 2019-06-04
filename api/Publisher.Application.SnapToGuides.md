---
title: Application.SnapToGuides property (Publisher)
keywords: vbapb10.chm131110
f1_keywords:
- vbapb10.chm131110
ms.prod: publisher
api_name:
- Publisher.Application.SnapToGuides
ms.assetid: 09894c02-3193-cd14-ff55-45920e461af9
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.SnapToGuides property (Publisher)

**True** for Microsoft Publisher to use the guides to align objects on a page in a publication. Read/write **Boolean**.


## Syntax

_expression_.**SnapToGuides**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

Boolean


## Example

This example adds horizontal and vertical ruler guides every half inch on the first page and then sets the options to align objects on the page to the guides.

```vb
Sub SetSnapOptions() 
 Dim intCount As Integer 
 Dim intPos As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To 16 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeVertical 
 Next 
 intPos = 0 
 For intCount = 1 To 21 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeHorizontal 
 Next 
 End With 
 With Application 
 .SnapToGuides = True 
 .SnapToObjects = True 
 End With 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]