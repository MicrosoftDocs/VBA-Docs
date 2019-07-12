---
title: SectionProperties.Delete method (PowerPoint)
keywords: vbapp10.chm725011
f1_keywords:
- vbapp10.chm725011
ms.prod: powerpoint
api_name:
- PowerPoint.SectionProperties.Delete
ms.assetid: 5a102ee7-60a1-64b1-db6c-6ba84447dd12
ms.date: 06/08/2017
localization_priority: Normal
---


# SectionProperties.Delete method (PowerPoint)

Deletes the section break that sets off the specified section, and optionally deletes all the slides in the section.


## Syntax

_expression_.**Delete**( `_sectionIndex_`, `_deleteSlides_` )

_expression_ A variable that represents a [SectionProperties](PowerPoint.SectionProperties.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _sectionIndex_|Required|**Integer**|The index of the section to delete.|
| _deleteSlides_|Required|**Boolean**|Whether to delete all the slides in the section.  **True**, to delete all the slides within the section; **False** not to delete them.|

## Remarks

If the presentation contains more than one section, you cannot delete the first section without deleting all of the slides in that section. 


## See also


[SectionProperties Object](PowerPoint.SectionProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]