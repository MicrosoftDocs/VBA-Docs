---
title: Pane.Pages property (Word)
keywords: vbawd10.chm157286419
f1_keywords:
- vbawd10.chm157286419
ms.prod: word
api_name:
- Word.Pane.Pages
ms.assetid: 18390c33-fd07-73a3-324f-37d09e1c99c6
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.Pages property (Word)

Returns a  **[Pages](Word.pages.md)** collection that represents the pages in a document.


## Syntax

_expression_.**Pages**

 _expression_ An expression that returns a '[Pane](Word.Pane.md)' object.


## Example

The following example creates a line 0.5 inch from the upper-left corner of the active document across the page to the lower-right corner of the page, 0.5 inch from the right and bottom edges of the page.


```vb
Dim objPage As Page 
 
Set objPage = ActiveDocument.ActiveWindow.Panes(1).Pages(1) 
 
'Add new line to document 
ActiveDocument.Shapes.AddLine _ 
 InchesToPoints(0.5), _ 
 InchesToPoints(0.5), _ 
 objPage.Width - InchesToPoints(0.5), _ 
 objPage.Height - InchesToPoints(0.5) 

```


## See also


[Pane Object](Word.Pane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]