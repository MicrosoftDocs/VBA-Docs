---
title: TextStyles.Add method (Publisher)
keywords: vbapb10.chm5898244
f1_keywords:
- vbapb10.chm5898244
ms.prod: publisher
api_name:
- Publisher.TextStyles.Add
ms.assetid: 56bb84a2-5632-1baa-4b97-3c48d43367bf
ms.date: 06/15/2019
localization_priority: Normal
---


# TextStyles.Add method (Publisher)

Adds a new **[TextStyle](Publisher.TextStyle.md)** object to the specified **TextStyles** collection and returns the new **TextStyle** object.


## Syntax

_expression_.**Add** (_StyleName_, _Font_, _ParagraphFormat_, _BasedOn_)

_expression_ A variable that represents a **[TextStyles](Publisher.TextStyles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_StyleName_|Required| **String**|The name of the new text style. If the name matches an existing text style, the existing text style is overwritten.|
|_Font_|Optional| **[Font](Publisher.Font.md)** |The font settings to apply to the new text style.|
|_ParagraphFormat_|Optional| **[ParagraphFormat](Publisher.ParagraphFormat.md)** |The paragraph formatting to apply to the new text style.|
|_BasedOn_|Optional| **String**|The name of the text style on which the new text style is based. If the name does not match an existing text style, an error occurs.|


## Return value

TextStyle


## Example

The following example adds a new text style to the active publication based on the Normal text style.

```vb
Dim tsNew As TextStyle 
 
Set tsNew = ActiveDocument.TextStyles _ 
 .Add(StyleName:="Title", BasedOn:="Normal")
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]