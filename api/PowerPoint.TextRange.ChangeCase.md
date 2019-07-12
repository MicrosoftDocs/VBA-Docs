---
title: TextRange.ChangeCase method (PowerPoint)
keywords: vbapp10.chm569031
f1_keywords:
- vbapp10.chm569031
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.ChangeCase
ms.assetid: a14edb26-7ec3-5fb5-7590-cd67a75c1f03
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.ChangeCase method (PowerPoint)

Changes the case of the specified text.


## Syntax

_expression_. `ChangeCase`( `_Type_` )

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[PpChangeCase](PowerPoint.PpChangeCase.md)**|Specifies the way the case will be changed.|

## Example

This example sets title case capitalization for the title on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes.Title.TextFrame _
    .TextRange.ChangeCase ppCaseTitle
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]