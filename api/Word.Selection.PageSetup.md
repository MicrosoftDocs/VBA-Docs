---
title: Selection.PageSetup property (Word)
keywords: vbawd10.chm158663757
f1_keywords:
- vbawd10.chm158663757
ms.prod: word
api_name:
- Word.Selection.PageSetup
ms.assetid: 65e8953b-0b52-b31f-1c81-e607a37ba916
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.PageSetup property (Word)

Returns a  **[PageSetup](Word.PageSetup.md)** object that's associated with the specified selection.


## Syntax

_expression_.**PageSetup**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example sets the header and footer distance to 18 points (0.25 inch) from the top and bottom of the page, respectively. This formatting change is applied to the section that contains the selection.


```vb
With Selection.PageSetup 
 .FooterDistance = 18 
 .HeaderDistance = 18 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]