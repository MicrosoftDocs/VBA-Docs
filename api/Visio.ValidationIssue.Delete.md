---
title: ValidationIssue.Delete method (Visio)
keywords: vis_sdr.chm18616165
f1_keywords:
- vis_sdr.chm18616165
ms.prod: visio
api_name:
- Visio.ValidationIssue.Delete
ms.assetid: a585713e-b394-5e5f-e5b2-259dacbe8bec
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationIssue.Delete method (Visio)

Deletes the  **[ValidationIssue](Visio.ValidationIssue.md)** object from the document.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[ValidationIssue](Visio.ValidationIssue.md)** object.


## Return value

 **Nothing**


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Delete** method to delete validation issues that belong to a particular validation rule set.


```vb
Set vsoDocument = Visio.ActiveDocument 
Set vsoIssues = vsoDocument.Validation.Issues
intIssueTotal = vsoIssues.Count
intIssueNumber = 1

' Iterate through the validation issues.
 For intCurrentIssue = 1 To intIssueTotal
      Set vsoIssue = vsoDocument.Validation.Issues(intIssueNumber)
      
     ' Delete the issues that belong to the vsoValidationRuleSet rule set.
     If vsoIssue.Rule.RuleSet Is vsoValidationRuleSet Then
         vsoIssue.Delete
     Else
        intIssueNumber = intIssueNumber + 1
     End If
     
 Next intCurrentIssue
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]