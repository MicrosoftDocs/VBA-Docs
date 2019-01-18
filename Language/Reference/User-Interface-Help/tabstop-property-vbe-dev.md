---
title: TabStop property (VBE.Dev)
ms.prod: office
ms.assetid: c1672383-72cf-4bb0-b1fa-96c830147f21
ms.date: 11/16/2018
localization_priority: Normal
---


# TabStop property (VBE.Dev)

Indicates whether an object can receive [focus](../../Glossary/vbe-glossary.md#focus) when the user tabs to it.

## Syntax

_object_.**TabStop** [= _Boolean_ ]

The **TabStop** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the object is a tab stop.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|Designates the object as a tab stop (default).|
|**False**|Bypasses the object when the user is tabbing, although the object still holds its place in the actual tab order, as determined by the **TabIndex** property.|

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]