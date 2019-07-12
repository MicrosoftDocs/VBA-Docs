---
title: Font.NameOther property (PowerPoint)
keywords: vbapp10.chm575019
f1_keywords:
- vbapp10.chm575019
ms.prod: powerpoint
api_name:
- PowerPoint.Font.NameOther
ms.assetid: 64f62838-635c-9b6d-082a-06fe698685e1
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.NameOther property (PowerPoint)

Returns or sets the font used for characters whose character set numbers are greater than 127. Read/write.


## Syntax

_expression_. `NameOther`

_expression_ A variable that represents a [Font](PowerPoint.Font.md) object.


## Return value

String


## Remarks

In the U.S. English version of Microsoft PowerPoint, this property is read-only and the default value is Times New Roman. Use the  **[Replace](PowerPoint.Fonts.Replace.md)** method to change a font in a presentation. The **NameOther** property setting is the same as the **NameASCII** property setting except when the **NameASCII** property is set to "Use FE Font."


## Example

This example sets the font used for characters whose character set numbers are greater than 127, for the first member of the  **Fonts** collection.


```vb
ActivePresentation.Fonts(1).NameOther = "Tahoma"
```


## See also


[Font Object](PowerPoint.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]