---
title: Style.ListLevelNumber property (Word)
keywords: vbawd10.chm153878543
f1_keywords:
- vbawd10.chm153878543
ms.prod: word
api_name:
- Word.Style.ListLevelNumber
ms.assetid: c237a4ab-71e2-d8e4-21a0-bc7c4c3c892a
ms.date: 06/08/2017
localization_priority: Normal
---


# Style.ListLevelNumber property (Word)

Returns the list level for the specified style. Read-only  **Long**.


## Syntax

_expression_. `ListLevelNumber`

_expression_ Required. A variable that represents a '[Style](Word.Style.md)' object.


## Example

This example displays the list level for the Heading 3 style.


```vb
Msgbox ActiveDocument.Styles(wdStyleHeading3).ListLevelNumber
```


## See also


[Style Object](Word.Style.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]