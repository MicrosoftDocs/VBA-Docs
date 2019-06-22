---
title: ValidationRules.Add method (Visio)
keywords: vis_sdr.chm18316005
f1_keywords:
- vis_sdr.chm18316005
ms.prod: visio
api_name:
- Visio.ValidationRules.Add
ms.assetid: 14b0ab24-5ff6-cde5-8311-ccf2989712c9
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationRules.Add method (Visio)

Adds a new, empty **[ValidationRule](Visio.ValidationRule.md)** object to the **ValidationRules** collection of the document.


## Syntax

_expression_.**Add** (_NameU_)

_expression_ A variable that represents a **[ValidationRules](Visio.ValidationRules.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NameU_|Required| **String**|The universal name to assign to the new validation rule.|

## Return value

 **ValidationRule**


## Remarks

If the  _NameU_ parameter is not a valid string, Visio returns an Invalid Parameter error.

The default property values of the new validation rule are as follows: 

- **[Category](Visio.ValidationRule.Category.md)** = [empty]
- **[Description](Visio.ValidationRule.Description.md)** = "Unknown"
- **[FilterExpression](Visio.ValidationRule.FilterExpression.md)** = [empty]
- **[Ignored](Visio.ValidationRule.Ignored.md)** = **False**
- **[TargetType](Visio.ValidationRule.TargetType.md)** = **visRuleTargetShape**
- **[TestExpression](Visio.ValidationRule.TestExpression.md)** = [empty]


## Example

The following sample code is based on code provided by: [David Parker](https://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the **Add** method to add a new validation rule named "UngluedConnector" to an existing validation rule set named "Fault Tree Analysis" in the active document.




```vb
Public Sub Add_Example()

    Dim vsoValidationRule As Visio.ValidationRule
    Dim vsoValidationRuleSet As Visio.ValidationRuleSet
    Dim strValidationRuleSetNameU As String
    Dim strValidationRuleNameU As String
    
    strValidationRuleSetNameU = "Fault Tree Analysis"
    strValidationRuleNameU = "UngluedConnector"
    
    Set vsoValidationRuleSet = ActiveDocument.Validation.RuleSets(strValidationRuleSetNameU)
    Set vsoValidationRule = vsoValidationRuleSet.Rules.Add(strValidationRuleNameU)

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]