---
title: TabStops.Add method (Publisher)
keywords: vbapb10.chm5570565
f1_keywords:
- vbapb10.chm5570565
ms.prod: publisher
api_name:
- Publisher.TabStops.Add
ms.assetid: 23536810-e851-c0ac-22e2-fab41582d612
ms.date: 06/15/2019
localization_priority: Normal
---


# TabStops.Add method (Publisher)

Adds a new tab stop to the specified **TabStops** collection.


## Syntax

_expression_.**Add** (_Position_, _Alignment_, _Leader_)

_expression_ A variable that represents a **[TabStops](Publisher.TabStops.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Position_|Required| **Variant**|The horizontal position of the new tab stop relative to the left edge of the text frame. Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings are evaluated in the units specified and can be in any measurement unit supported by Microsoft Publisher (for example, "2.5 in").|
|_Alignment_|Required| **[PbTabAlignmentType](publisher.pbtabalignmenttype.md)**|The alignment setting for the tab stop. Can be one of the **PbTabAlignmentType** constants.|
|_Leader_|Required| **[PbTabLeaderType](publisher.pbtableadertype.md)**|The type of leader for the tab stop. Can be one of the **PbTabLeaderType** constants.|


## Example

The following example adds a new left-aligned tab stop 0.5 inches from the left edge of the specified text frame.

```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Tabs _ 
 .Add Position:="0.5 in", _ 
 Alignment:=pbTabAlignmentLeading, _ 
 Leader:=pbTabLeaderNone
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]