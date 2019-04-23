---
title: OldHeight, OldWidth properties
keywords: fm20.chm2001620
f1_keywords:
- fm20.chm2001620
ms.prod: office
ms.assetid: cd2c0dfb-85f3-2381-128b-4d964829e7b0
ms.date: 11/16/2018
localization_priority: Normal
---


# OldHeight, OldWidth properties

Returns the previous height or width, in [points](../../Glossary/vbe-glossary.md#point), of the control.

## Syntax

_object_.**OldHeight** <br/>
_object_.**OldWidth**
 
The **OldHeight** and **OldWidth** property syntaxes have these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

**OldHeight** and **OldWidth** are read-only.

The **OldHeight** and **OldWidth** properties are automatically updated when you move or size a control. If you change the size of a control, the **Height** and **Width** properties store the new height, and **OldHeight** and **OldWidth** store the previous height.

These properties are valid only in the [Layout](layout-event.md) event.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]