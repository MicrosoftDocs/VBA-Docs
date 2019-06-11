---
title: PhoneticGuide.Raise property (Publisher)
keywords: vbapb10.chm6160389
f1_keywords:
- vbapb10.chm6160389
ms.prod: publisher
api_name:
- Publisher.PhoneticGuide.Raise
ms.assetid: 8c7bd7e9-1b63-ded0-5021-99995296ad08
ms.date: 06/12/2019
localization_priority: Normal
---


# PhoneticGuide.Raise property (Publisher)

Returns a **Variant** indicating the distance between the top of the base text and the bottom of the guide text. Read-only.


## Syntax

_expression_.**Raise**

_expression_ A variable that represents a **[PhoneticGuide](Publisher.PhoneticGuide.md)** object.


## Return value

Variant


## Remarks

Numeric set values are in [points](../language/glossary/vbe-glossary.md#point); strings can be any measurement unit supported by Microsoft Publisher. Return values are always in points.


## Example

The following example places the phonetic guide for shape one in the active publication five points above the base text.

```vb
Dim phoGuide As PhoneticGuide 
 
Set phoGuide = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Fields(1).PhoneticGuide 
 
With phoGuide 
 .Raise = 5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]