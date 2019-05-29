---
title: PageSetup.HeaderDistance property (Word)
keywords: vbawd10.chm158400624
f1_keywords:
- vbawd10.chm158400624
ms.prod: word
api_name:
- Word.PageSetup.HeaderDistance
ms.assetid: fee422f6-ecf0-0470-2845-b8694636a76e
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.HeaderDistance property (Word)

Returns or sets the distance (in points) between the header and the top of the page. Read/write  **Single**.


## Syntax

_expression_. `HeaderDistance`

_expression_ A variable that represents a **[PageSetup](Word.PageSetup.md)** object.


## Example

This example displays the distance between the header and the top of the page. The  **[PointsToInches](Word.Global.PointsToInches.md)** method is used to convert points to inches.


```vb
Dim sngDistance As Single 
 
sngDistance = ActiveDocument.PageSetup.HeaderDistance 
Msgbox PointsToInches(sngDistance) & " inches"
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]