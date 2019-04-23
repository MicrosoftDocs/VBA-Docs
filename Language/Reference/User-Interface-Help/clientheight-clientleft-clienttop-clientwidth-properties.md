---
title: ClientHeight, ClientLeft, ClientTop, ClientWidth properties
keywords: fm20.chm2000910
f1_keywords:
- fm20.chm2000910
ms.prod: office
ms.assetid: d0754b52-156b-f8a4-3b28-9ce3020bc5f7
ms.date: 11/15/2018
localization_priority: Normal
---


# ClientHeight, ClientLeft, ClientTop, ClientWidth properties

Define the dimensions and location of the display area of a **[TabStrip](tabstrip-control.md)**.

## Syntax

_object_.**ClientHeight** [ = _Single_ ] <br/>
_object_.**ClientLeft** [ = _Single_ ] <br/>
_object_.**ClientTop** [ = _Single_ ] <br/>
_object_.**ClientWidth** [ = _Single_ ] <br/>

The **ClientHeight**, **ClientLeft**, **ClientTop**, and **ClientWidth** property syntaxes have these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Single_|Optional. For **ClientHeight** and **ClientWidth**, specifies the height or width, in points, of the display area. For **ClientLeft** and **ClientTop**, specifies the distance, in points, from the top or left edge of the **TabStrip** container.|

## Remarks

At [run time](../../Glossary/vbe-glossary.md#run-time), **ClientLeft**, **ClientTop**, **ClientHeight**, and **ClientWidth** automatically store the coordinates and dimensions of the **TabStrip** internal area, which is shared by objects in the **TabStrip**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]