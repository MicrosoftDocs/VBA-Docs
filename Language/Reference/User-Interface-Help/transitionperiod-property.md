---
title: TransitionPeriod property
keywords: fm20.chm5225109
f1_keywords:
- fm20.chm5225109
ms.prod: office
api_name:
- Office.TransitionPeriod
ms.assetid: cfdda5c3-244b-4404-d6a8-544755056473
ms.date: 11/16/2018
localization_priority: Normal
---


# TransitionPeriod property

Specifies the duration, in milliseconds, of a transition effect.

## Syntax

_object_.**TransitionPeriod** [= _Long_ ]

The **TransitionPeriod** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. How long it takes to complete the transition from one page to another.|

## Remarks

Any integer from zero to 10000 is a valid setting for this property. Setting the **TransitionPeriod** property to zero disables the transition effect; setting **TransitionPeriod** to 10000 creates a 10-second transition.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]