---
title: Global.PortraitFontNames property (Word)
keywords: vbawd10.chm163119117
f1_keywords:
- vbawd10.chm163119117
ms.prod: word
api_name:
- Word.Global.PortraitFontNames
ms.assetid: 07627cb8-a47f-14c9-b630-de9318e9e3d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.PortraitFontNames property (Word)

Returns a  **FontNames** object that includes the names of all the available portrait fonts.


## Syntax

_expression_. `PortraitFontNames`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]