---
title: Selection.HeaderFooter property (Word)
keywords: vbawd10.chm158662962
f1_keywords:
- vbawd10.chm158662962
ms.prod: word
api_name:
- Word.Selection.HeaderFooter
ms.assetid: b2eeeb83-49bf-236e-e795-6231ff20e368
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.HeaderFooter property (Word)

Returns a  **[HeaderFooter](Word.HeaderFooter.md)** object for the specified selection. Read-only.


## Syntax

 _expression_. `HeaderFooter`

 _expression_ A variable that represents a '[Selection](Word.Selection.md)' object.


## Remarks

An error occurs if the selection isn't located within a header or footer.


## Example

This example adds a centered page number to the current page footer.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageFooter 
End With 
Selection.HeaderFooter.PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberCenter
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]