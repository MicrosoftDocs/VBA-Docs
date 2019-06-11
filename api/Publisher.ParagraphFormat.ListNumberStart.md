---
title: ParagraphFormat.ListNumberStart property (Publisher)
keywords: vbapb10.chm5439527
f1_keywords:
- vbapb10.chm5439527
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListNumberStart
ms.assetid: 8e17fdaa-f53e-26c4-d92b-8ead65c28555
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.ListNumberStart property (Publisher)

Sets or retrieves a **Long** that represents the starting number of a list. Read/write.


## Syntax

_expression_.**ListNumberStart**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Long


## Remarks

Returns an "Access Denied" message if the list is not a numbered list.


## Example

This example sets the list type of a **ParagraphFormat** object to **pbListTypeArabic** (**[PbListType](publisher.pblisttype.md)** enumeration) and sets the **ListNumber** property to 4.

```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
 With objParaForm 
 .SetListType pbListTypeArabic 
 .ListNumberStart = 4 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]