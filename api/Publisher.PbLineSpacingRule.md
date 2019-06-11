---
title: PbLineSpacingRule enumeration (Publisher)
keywords: vbapb10.chm65552
f1_keywords:
- vbapb10.chm65552
ms.prod: publisher
api_name:
- Publisher.PbLineSpacingRule
ms.assetid: 64a5742e-b361-8e9a-31e4-8953b23ded14
ms.date: 06/12/2019
localization_priority: Normal
---


# PbLineSpacingRule enumeration (Publisher)

Represents the line spacing for the specified paragraphs.

<br/>

|Name|Value|Description|
|:-----|:-----|:-----|
| **pbLineSpacing1pt5**|1|Sets the spacing for specified paragraphs to one-and-a-half lines.|
| **pbLineSpacingDouble**|2|Double-spaces the specified paragraphs (sets paragraph line spacing to two lines).|
| **pbLineSpacingExactly**|4|Sets the line spacing to exactly the value specified in the _Spacing_ argument, even if a larger font is used within the paragraph.|
| **pbLineSpacingMixed**|-9999999|A return value for the **[LineSpacing](Publisher.ParagraphFormat.LineSpacing.md)** property that indicates that line spacing is a combination of values for the specified paragraphs.|
| **pbLineSpacingMultiple**|5|Sets the line spacing to the value specified in the _Spacing_ argument.|
| **pbLineSpacingSingle**|0|Single spaces the specified paragraphs (sets paragraph line spacing to one space).|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]