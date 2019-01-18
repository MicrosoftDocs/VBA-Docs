---
title: SpinDown, SpinUp events
keywords: fm20.chm2000220
f1_keywords:
- fm20.chm2000220
ms.prod: office
ms.assetid: 4e6e4395-1622-eb97-59d0-2b52a22d6528
ms.date: 11/15/2018
localization_priority: Normal
---


# SpinDown, SpinUp events

SpinDown occurs when the user clicks the lower or left spin-button arrow. SpinUp occurs when the user clicks the upper or right spin-button arrow.

## Syntax

**Private Sub**_object_ _**SpinDown( )** <br/>
**Private Sub**_object_ _**SpinUp( )**

The **SpinDown** and **SpinUp** event syntaxes have these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

The SpinDown event decreases the **Value** property. The SpinUp event increases **Value**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]