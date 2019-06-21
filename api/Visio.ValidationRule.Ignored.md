---
title: ValidationRule.Ignored property (Visio)
keywords: vis_sdr.chm18462650
f1_keywords:
- vis_sdr.chm18462650
ms.prod: visio
api_name:
- Visio.ValidationRule.Ignored
ms.assetid: e99a629b-f3de-fbd0-82d9-e821d18500c3
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationRule.Ignored property (Visio)

Determines whether the validation rule is currently ignored. Read/write.


## Syntax

_expression_. `Ignored`

_expression_ A variable that represents a **[ValidationRule](Visio.ValidationRule.md)** object.


## Return value

 **Boolean**


## Remarks

Issues that pertain to an ignored rule are still recorded but, by default, they are not displayed in the  **Issues** window.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Ignored** property to specify that the validation rule named "Unglued2DShape" should not be ignored.


```vb
Set vsoValidationRule = vsoValidationRuleSet.Rules.Add("Unglued2DShape")
vsoValidationRule.Ignored = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]