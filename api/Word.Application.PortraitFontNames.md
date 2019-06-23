---
title: Application.PortraitFontNames property (Word)
keywords: vbawd10.chm158334989
f1_keywords:
- vbawd10.chm158334989
ms.prod: word
api_name:
- Word.Application.PortraitFontNames
ms.assetid: 21c3802b-43ad-3d8f-34de-af9af4d29bcf
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PortraitFontNames property (Word)

Returns a  **[FontNames](Word.FontNames.md)** object that includes the names of all the available portrait fonts.


## Syntax

_expression_. `PortraitFontNames`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example inserts a list of portrait fonts at the insertion point.


```vb
For Each aFont In PortraitFontNames 
 With Selection 
 .Collapse Direction:=wdCollapseEnd 
 .InsertAfter aFont 
 .InsertParagraphAfter 
 .Collapse Direction:=wdCollapseEnd 
 End With 
Next aFont
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]