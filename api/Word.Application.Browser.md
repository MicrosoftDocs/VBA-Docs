---
title: Application.Browser property (Word)
keywords: vbawd10.chm158334992
f1_keywords:
- vbawd10.chm158334992
ms.prod: word
api_name:
- Word.Application.Browser
ms.assetid: 79b1967d-e661-8953-7bb2-a35eadbfae54
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Browser property (Word)

Returns a  **[Browser](Word.Browser.md)** object that represents the **Select Browse Object** tool on the vertical scroll bar. Read-only.


## Syntax

 _expression_. `Browser`

 _expression_ A variable that represents an '[Application](Word.Application.md)' object.


## Example

This example moves to the next footnote reference mark in the active document.


```vb
With Application.Browser 
 .Target = wdBrowseFootnote 
 .Next 
End With
```

This example moves to the next field in the active document. The text from the initial selection to the next field is formatted as bold.




```vb
Selection.ExtendMode = True 
With Application.Browser 
 .Target = wdBrowseField 
 .Next 
End With 
With Selection 
 .Font.Bold = True 
 .ExtendMode = False 
 .Collapse Direction:=wdCollapseEnd 
End With
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]