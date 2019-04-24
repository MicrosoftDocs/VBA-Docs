---
title: Paragraph.Style property (Word)
keywords: vbawd10.chm156696676
f1_keywords:
- vbawd10.chm156696676
ms.prod: word
api_name:
- Word.Paragraph.Style
ms.assetid: a6ac7009-4018-b873-8db5-6c86afd11a22
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Style property (Word)

Returns or sets the style for the specified object. Read/write  **Variant**.


## Syntax

_expression_.**Style**

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

To set this property, specify the local name of the style, an integer, a  **[WdBuiltinStyle](Word.WdBuiltinStyle.md)** constant, or an object that represents the style.


## Example

This example displays the style for each paragraph in the active document.


```vb
For Each para in ActiveDocument.Paragraphs 
 MsgBox para.Style 
Next para
```

This example sets alternating styles of Heading 3 and Normal for all the paragraphs in the active document.




```vb
For i = 1 To ActiveDocument.Paragraphs.Count 
 If i Mod 2 = 0 Then 
 ActiveDocument.Paragraphs(i).Style = wdStyleNormal 
 Else: ActiveDocument.Paragraphs(i).Style = wdStyleHeading3 
 End If 
Next i
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
