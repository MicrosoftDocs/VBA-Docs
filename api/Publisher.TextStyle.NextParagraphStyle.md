---
title: TextStyle.NextParagraphStyle property (Publisher)
keywords: vbapb10.chm5963784
f1_keywords:
- vbapb10.chm5963784
ms.prod: publisher
api_name:
- Publisher.TextStyle.NextParagraphStyle
ms.assetid: 2b31b883-c26d-3be8-7145-f8e3cf1ba5cc
ms.date: 06/15/2019
localization_priority: Normal
---


# TextStyle.NextParagraphStyle property (Publisher)

Returns or sets a **String** that represents the paragraph style that follows the specified text style when a user presses Enter. Read/write.


## Syntax

_expression_.**NextParagraphStyle**

_expression_ A variable that represents a **[TextStyle](Publisher.TextStyle.md)** object.


## Return value

String


## Example

This example creates a new text style and specifies that the text style following the new text style is the Normal style.

```vb
Sub CreateNewTextStyle() 
 Dim styNew As TextStyle 
 Dim fntStyle As Font 
 
 Set styNew = ActiveDocument.TextStyles.Add(StyleName:="Heading 1") 
 Set fntStyle = styNew.Font 
 
 With fntStyle 
 .Name = "Tahoma" 
 .Bold = msoTrue 
 .Size = 15 
 End With 
 
 With styNew 
 .Font = fntStyle 
 .NextParagraphStyle = "Normal" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]