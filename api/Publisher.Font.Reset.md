---
title: Font.Reset method (Publisher)
keywords: vbapb10.chm5373993
f1_keywords:
- vbapb10.chm5373993
ms.prod: publisher
api_name:
- Publisher.Font.Reset
ms.assetid: 7a81d7f9-4db9-3ce1-188d-2b4719b57fff
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.Reset method (Publisher)

Removes manual paragraph or text formatting from the specified object and leaves only the formatting specified by the current text style.


## Syntax

_expression_.**Reset**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


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