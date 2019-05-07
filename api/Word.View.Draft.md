---
title: View.Draft property (Word)
keywords: vbawd10.chm161808386
f1_keywords:
- vbawd10.chm161808386
ms.prod: word
api_name:
- Word.View.Draft
ms.assetid: 9a0dd1df-6d5d-babc-02f8-74bf7e651226
ms.date: 06/08/2017
localization_priority: Normal
---


# View.Draft property (Word)

 **True** if all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display. Read/write **Boolean**.


## Syntax

_expression_.**Draft**

_expression_ A variable that represents a '[View](Word.View.md)' object.


## Example

This example displays the contents of the window for Document1 in the draft font.


```vb
Windows("Document1").View.Draft = True
```

This example toggles the draft font option for the active window.




```vb
ActiveDocument.ActiveWindow.View.Draft = _ 
 Not ActiveDocument.ActiveWindow.View.Draft
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]