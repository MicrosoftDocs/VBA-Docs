---
title: Fonts.Replace method (PowerPoint)
keywords: vbapp10.chm528004
f1_keywords:
- vbapp10.chm528004
ms.prod: powerpoint
api_name:
- PowerPoint.Fonts.Replace
ms.assetid: 666bcfad-b87e-b63b-70c1-ca0873cf9f94
ms.date: 06/08/2017
localization_priority: Normal
---


# Fonts.Replace method (PowerPoint)

Replaces a font in the  **Fonts** collection.


## Syntax

_expression_.**Replace** (_Original_, _Replacement_)

_expression_ A variable that represents a [Fonts](PowerPoint.Fonts.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Original_|Required|**String**|The name of the font to replace.|
| _Replacement_|Required|**String**|The name of the replacement font.|

## Example

This example replaces the Times New Roman font with the Courier font in the active presentation.


```vb
Application.ActivePresentation.Fonts _
    .Replace Original:="Times New Roman", Replacement:="Courier"
```


## See also


[Fonts Object](PowerPoint.Fonts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]