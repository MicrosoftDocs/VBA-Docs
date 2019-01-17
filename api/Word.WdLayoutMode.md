---
title: WdLayoutMode enumeration (Word)
ms.prod: word
api_name:
- Word.WdLayoutMode
ms.assetid: bdcb65a5-c198-c515-d6a6-25e43e46b70f
ms.date: 06/08/2017
localization_priority: Normal
---


# WdLayoutMode enumeration (Word)

Specifies how text is laid out in the layout mode for the current document.



|Name|Value|Description|
|:-----|:-----|:-----|
| **wdLayoutModeDefault**|0|No grid is used to lay out text.|
| **wdLayoutModeGenko**|3|Text is laid out on a grid; the user specifies the number of lines and the number of characters per line. As the user types, Microsoft Word automatically aligns characters with gridlines.|
| **wdLayoutModeGrid**|1|Text is laid out on a grid; the user specifies the number of lines and the number of characters per line. As the user types, Microsoft Word doesn't automatically align characters with gridlines.|
| **wdLayoutModeLineGrid**|2|Text is laid out on a grid; the user specifies the number of lines, but not the number of characters per line.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]