---
title: SectionProperties.SectionID method (PowerPoint)
keywords: vbapp10.chm725012
f1_keywords:
- vbapp10.chm725012
ms.prod: powerpoint
api_name:
- PowerPoint.SectionProperties.SectionID
ms.assetid: eec3a584-8f97-ae9f-9630-0b34964a5c02
ms.date: 06/08/2017
localization_priority: Normal
---


# SectionProperties.SectionID method (PowerPoint)

Returns a string that represents the unique identifier of the specified section.


## Syntax

_expression_. `SectionID`( `_sectionIndex_` )

_expression_ A variable that represents a [SectionProperties](PowerPoint.SectionProperties.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _sectionIndex_|Required|**Integer**|The index of the section.|

## Return value

String


## Remarks

The returned string is in the form "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXX}", where each "X" is a hexadecimal digid.


## See also


[SectionProperties Object](PowerPoint.SectionProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]