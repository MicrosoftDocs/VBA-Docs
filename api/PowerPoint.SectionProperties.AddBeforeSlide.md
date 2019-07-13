---
title: SectionProperties.AddBeforeSlide method (PowerPoint)
keywords: vbapp10.chm725008
f1_keywords:
- vbapp10.chm725008
ms.prod: powerpoint
api_name:
- PowerPoint.SectionProperties.AddBeforeSlide
ms.assetid: ad11901c-3e64-7c08-ae89-a1285a6fa075
ms.date: 06/08/2017
localization_priority: Normal
---


# SectionProperties.AddBeforeSlide method (PowerPoint)

Adds a section immediately before the specified slide index, and returns the index of the new section.


## Syntax

_expression_. `AddBeforeSlide`( `_SlideIndex_`, `_sectionName_` )

_expression_ A variable that represents a [SectionProperties](PowerPoint.SectionProperties.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SlideIndex_|Required|**Integer**|The index of the slide before which to add the section.|
| _sectionName_|Required|**String**|The name of the new section.|

## Return value

Integer


## Remarks

The indices of sections after the newly inserted section are automatically incremented by one.

If a section break exists immediately before the specified slide index, the new section is placed after the section break, with the result that the preceding section is now empty, and the specified slide index is now the first slide of the new section.

If the presentation does not contain any sections and you call this method, passing a SlideIndex value greater than 1, a new section is created before the first slide and given the default section name.


## See also


[SectionProperties Object](PowerPoint.SectionProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]