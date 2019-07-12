---
title: PrintOptions.PrintFontsAsGraphics property (PowerPoint)
keywords: vbapp10.chm517013
f1_keywords:
- vbapp10.chm517013
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.PrintFontsAsGraphics
ms.assetid: f782be2c-9787-72e3-139e-163041e066f7
ms.date: 06/08/2017
localization_priority: Normal
---


# PrintOptions.PrintFontsAsGraphics property (PowerPoint)

Determines whether TrueType fonts are printed as graphics. Read/write.


## Syntax

_expression_. `PrintFontsAsGraphics`

_expression_ A variable that represents a [PrintOptions](PowerPoint.PrintOptions.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **PrintFontsAsGraphics** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**| TrueType fonts are not printed as graphics.|
|**msoTrue**| TrueType fonts are printed as graphics.|

## Example

This example specifies that TrueType fonts in the active presentation be printed as graphics.


```vb
ActivePresentation.PrintOptions.PrintFontsAsGraphics = msoTrue
```


## See also


[PrintOptions Object](PowerPoint.PrintOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]