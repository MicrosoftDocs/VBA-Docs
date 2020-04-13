---
title: Selection.Font property (Word)
keywords: vbawd10.chm158662661
f1_keywords:
- vbawd10.chm158662661
ms.prod: word
api_name:
- Word.Selection.Font
ms.assetid: c2a24190-62fa-09c4-7c47-90a7ecf20d97
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Font property (Word)

Returns or sets a  **[Font](Word.Font.md)** object that represents the character formatting of the specified object. Read/write.


## Syntax

_expression_.**Font**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

To set the **Font** property, specify an expression that returns a **Font** object.


## Example

This example displays the font of the selected text.


```vb
MsgBox Selection.Font.Name
```

This example applies the character formatting of the selected text to the first paragraph in the active document.




```vb
Set myFont = Selection.Font.Duplicate 
ActiveDocument.Paragraphs(1).Range.Font = myFont
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]