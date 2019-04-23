---
title: Page.Breaks property (Word)
keywords: vbawd10.chm11075591
f1_keywords:
- vbawd10.chm11075591
ms.prod: word
api_name:
- Word.Page.Breaks
ms.assetid: 13aed7c7-cf67-1456-7842-d113dfc00b31
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Breaks property (Word)

Returns a  **Breaks** collection that represents the breaks on a page. .


## Syntax

_expression_. `Breaks`

_expression_ Required. A variable that represents a '[Page](Word.Page.md)' object.


## Remarks

The **Breaks** collection includes page, column, and section breaks. Use the **Breaks** collection and the related objects and properties to programmatically define page layout in a document.


## Example

The following example returns the breaks on the first page in the active document.


```vb
Dim objBreaks As Breaks 
 
Set objBreaks = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Breaks
```


## See also


[Page Object](Word.Page.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]