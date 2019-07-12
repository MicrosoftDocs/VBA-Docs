---
title: Presentation.PublishSlides method (PowerPoint)
keywords: vbapp10.chm583108
f1_keywords:
- vbapp10.chm583108
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PublishSlides
ms.assetid: 2f5c569a-fc4d-01ae-eae7-f1894541e08e
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.PublishSlides method (PowerPoint)

Creates a Web presentation (in HTML format) containing slides from any loaded presentation. You can view the published presentation in a web browser.


## Syntax

_expression_. `PublishSlides`( `_SlideLibraryUrl_`, `_Overwrite_` )

 _expression_ An expression that returns a [Presentation](PowerPoint.Presentation.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SlideLibraryUrl_|Required|**String**|URL to the slide library.|
| _Overwrite_|Optional|**Boolean**|**True** if the original presentation should be overwritten.|

## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]