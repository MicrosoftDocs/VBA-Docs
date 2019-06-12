---
title: ParagraphFormat.LineSpacingRule property (Publisher)
keywords: vbapb10.chm5439505
f1_keywords:
- vbapb10.chm5439505
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.LineSpacingRule
ms.assetid: e9855daa-59f4-a4b6-f153-5de515261414
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.LineSpacingRule property (Publisher)

Returns or sets a **[PbLineSpacingRule](Publisher.PbLineSpacingRule.md)** constant that represents the line spacing for the specified paragraphs. Read/write.


## Syntax

_expression_.**LineSpacingRule**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

PbLineSpacingRule


## Remarks

The **LineSpacingRule** property value can be one of the **PbLineSpacingRule** constants declared in the Microsoft Publisher type library.


## Example

This example formats the paragraph at the cursor position to double spacing.

```vb
Sub SetLineSpacing() 
 Selection.TextRange.ParagraphFormat 
 .LineSpacingRule = pbLineSpacingDouble 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]