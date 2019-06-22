---
title: ValidationRuleSet.Delete method (Visio)
keywords: vis_sdr.chm18216165
f1_keywords:
- vis_sdr.chm18216165
ms.prod: visio
api_name:
- Visio.ValidationRuleSet.Delete
ms.assetid: bd5fcd79-6cc6-7e24-b35f-944f9dee2cab
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationRuleSet.Delete method (Visio)

Deletes the **ValidationRuleSet** object from the document.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[ValidationRuleSet](Visio.ValidationRuleSet.md)** object.


## Return value

**Nothing**


## Remarks

Calling the **Delete** method also deletes all **[ValidationRule](Visio.ValidationRule.md)** objects that are associated with the validation rule set.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the **Delete** method to delete a validation rule set named Fault Tree Analysis from the active document.

```vb
' Delete a rule set from the active document.
Public Sub Delete_Example()

    Dim strValidationRuleSetNameU As String
    strValidationRuleSetNameU = "Fault Tree Analysis"
    
    ActiveDocument.Validation.RuleSets(strValidationRuleSetNameU).Delete
   
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]