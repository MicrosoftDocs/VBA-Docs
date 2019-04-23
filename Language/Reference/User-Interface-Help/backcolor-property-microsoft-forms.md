---
title: BackColor property (Microsoft Forms)
keywords: fm20.chm2000770
f1_keywords:
- fm20.chm2000770
ms.prod: office
ms.assetid: 70549eaf-d785-67e7-3f04-76151864d850
ms.date: 11/15/2018
localization_priority: Normal
---


# BackColor property (Microsoft Forms)

Specifies the [background color](../../Glossary/glossary-vba.md#background-color) of the object.

## Syntax

_object_.**BackColor** [= _Long_ ]

The **BackColor** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. A value or constant that determines the background color of an object.|

## Settings

You can use any integer that represents a valid color. You can also specify a color by using the [RGB](../../Glossary/glossary-vba.md#rgb) function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75.

## Remarks

You can only see the background color of an object if the **BackStyle** property is set to **fmBackStyleOpaque**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]