---
title: ValidationRuleSet.RuleSetFlags property (Visio)
keywords: vis_sdr.chm18262640
f1_keywords:
- vis_sdr.chm18262640
ms.prod: visio
api_name:
- Visio.ValidationRuleSet.RuleSetFlags
ms.assetid: fefa08cb-65d5-f4b2-619a-d6345cfd83f4
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationRuleSet.RuleSetFlags property (Visio)

Gets or sets special rule-set properties. Read/write.


## Syntax

_expression_. `RuleSetFlags`

_expression_ A variable that represents a **[ValidationRuleSet](Visio.ValidationRuleSet.md)** object.


## Return value

 **[VisRuleSetFlags](Visio.VisRuleSetFlags.md)**


## Remarks

The  **RuleSetFlags** property value must be one of the following **VisRuleSetFlags** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visRuleSetDefault**|0|The default set of rule-set properties. The rule set appears in the  **Rules to Check** list (click the **Check Diagram** arrow on the **Process** tab).|
| **visRuleSetHidden**|1|The rule set does not appear in the  **Rules to Check** list.|

## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **RuleSetFlags** property to set the properties for a validation rule set named "Connectivity" in the active document.


```vb
Set vsoDocument = Visio.ActiveDocument
Set vsoValidationRuleSet = vsoDocument.Validation.RuleSets.Add("Connectivity")
vsoValidationRuleSet.RuleSetFlags = Visio.VisRuleSetFlags.visRuleSetDefault
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]