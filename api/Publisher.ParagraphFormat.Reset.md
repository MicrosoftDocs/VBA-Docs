---
title: ParagraphFormat.Reset method (Publisher)
keywords: vbapb10.chm5439509
f1_keywords:
- vbapb10.chm5439509
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.Reset
ms.assetid: 8ef5c799-cace-133c-33d3-3454df2c2f24
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.Reset method (Publisher)

Removes manual paragraph or text formatting from the specified object and leaves only the formatting specified by the current text style.


## Syntax

_expression_.**Reset**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Nothing


## Example

The following example resets the character formatting of the text in shape one on page one of the active publication to the default character formatting for the current text style.

```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font.Reset
```

<br/>

The following example resets the paragraph formatting of the text in shape one on page one of the active publication to the default paragraph formatting for the current text style.

```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat.Reset
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]