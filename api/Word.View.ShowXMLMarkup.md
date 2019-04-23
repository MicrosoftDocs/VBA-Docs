---
title: View.ShowXMLMarkup property (Word)
keywords: vbawd10.chm161808430
f1_keywords:
- vbawd10.chm161808430
ms.prod: word
api_name:
- Word.View.ShowXMLMarkup
ms.assetid: 70873416-6ca8-18c7-550f-46973a7b0f6e
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowXMLMarkup property (Word)

Returns a  **Long** that represents whether XML tags are visible in a document.


## Syntax

_expression_. `ShowXMLMarkup`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Remarks

 **True** indicates that tags are visible. **False** indicates that tags are hidden. **wdToggle** allows you to switch between showing and hiding the XML markup.


## Example

The following example switches between showing and hiding the XML markup in the active document.


```vb
ActiveDocument.ActiveWindow.View.ShowXMLMarkup = wdToggle
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]