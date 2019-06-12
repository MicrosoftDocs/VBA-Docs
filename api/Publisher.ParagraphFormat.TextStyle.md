---
title: ParagraphFormat.TextStyle property (Publisher)
keywords: vbapb10.chm5439508
f1_keywords:
- vbapb10.chm5439508
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.TextStyle
ms.assetid: 8495c9c8-387e-a2e8-26cb-08f660dde985
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.TextStyle property (Publisher)

Returns or sets a **Variant** that represents the text style applied to a paragraph. Read/write.


## Syntax

_expression_.**TextStyle**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Variant


## Example

This example changes the text style of the selection if the selection isn't formatted with the Normal text style. This example assumes that text is selected in the active publication.

```vb
Sub SetTextStyle() 
 With Selection.TextRange.ParagraphFormat 
 If .TextStyle <> "Normal" Then _ 
 .TextStyle = "Normal" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]