---
title: XMLNode.PlaceholderText property (Word)
keywords: vbawd10.chm37748761
f1_keywords:
- vbawd10.chm37748761
ms.prod: word
api_name:
- Word.XMLNode.PlaceholderText
ms.assetid: a7c8fc01-ecb7-3587-8ae1-3c340446a304
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.PlaceholderText property (Word)

Sets or returns a  **String** that represents the text displayed for an element that contains no text.


## Syntax

_expression_. `PlaceholderText`

 _expression_ An expression that returns an '[XMLNode](Word.XMLNode.md)' object.


## Remarks

Placeholder text is displayed in Microsoft Word only when the **Show XML tags in the document** check box in the **XML Structure** task pane is cleared. The **Show XML tags in the document** check box corresponds to the **[ShowXMLMarkup](Word.View.ShowXMLMarkup.md)** property.


## Example

The following example inserts a new element into the active document at the insertion point and sets what text to display when tags are not displayed in the document.


```vb
Dim objNode As XMLNode 
 
Set objNode = Selection.XMLNodes.Add("catalog", "book") 
 
objNode.PlaceholderText = "Enter Book Information Here"
```


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]