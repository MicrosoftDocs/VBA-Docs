---
title: ValidationIssues.Clear method (Visio)
keywords: vis_sdr.chm18562410
f1_keywords:
- vis_sdr.chm18562410
ms.prod: visio
api_name:
- Visio.ValidationIssues.Clear
ms.assetid: e3792e98-a47e-2ce2-e1ff-995ccbf645eb
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationIssues.Clear method (Visio)

Removes all  **[ValidationIssue](Visio.ValidationIssue.md)** objects from the **[ValidationRules](Visio.ValidationRules.md)** collection of the document.


## Syntax

_expression_.**Clear**

_expression_ A variable that represents a **[ValidationIssues](Visio.ValidationIssues.md)** object.


## Return value

 **Nothing**


## Remarks

Calling the  **Clear** method also resets the **[Validation.LastValidatedDate](Visio.Validation.LastValidatedDate.md)** property value to 0 (zero).


## Example

The following sample code is based on code provided by: [David Parker](https://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the  **Clear** method to clear all validation issues from the active document.




```vb

Public Sub Clear_Example()

    ActiveDocument.Validation.Issues.Clear
    
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]