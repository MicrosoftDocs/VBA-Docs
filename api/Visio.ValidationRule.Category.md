---
title: ValidationRule.Category property (Visio)
keywords: vis_sdr.chm18413175
f1_keywords:
- vis_sdr.chm18413175
ms.prod: visio
api_name:
- Visio.ValidationRule.Category
ms.assetid: 2ceb2edc-26a0-7fe4-ba48-a07f6e922af1
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationRule.Category property (Visio)

Represents the text displayed in the  **Category** column of the **Issues** window. Read/write.


## Syntax

_expression_.**Category**

_expression_ A variable that represents a **[ValidationRule](Visio.ValidationRule.md)** object.


## Return value

 **String**


## Remarks

The length of the string assigned to the  **Category** property cannot exceed 255 characters.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Category** property to set the text of the **Category** column of the **Issues** window for a validation rule named "Unglued2DShape".


```vb
Set vsoValidationRule = vsoValidationRuleSet.Rules.Add("Unglued2DShape")
vsoValidationRule.Category = "Shapes"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]