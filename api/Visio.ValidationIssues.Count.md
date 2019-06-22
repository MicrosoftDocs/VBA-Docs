---
title: ValidationIssues.Count property (Visio)
keywords: vis_sdr.chm18513330
f1_keywords:
- vis_sdr.chm18513330
ms.prod: visio
api_name:
- Visio.ValidationIssues.Count
ms.assetid: 7077d75d-640c-32ee-fdf3-1be37407ab94
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationIssues.Count property (Visio)

Returns the number of  **[ValidationIssue](Visio.ValidationIssue.md)** objects in the collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[ValidationIssues](Visio.ValidationIssues.md)** object.


## Return value

 **Long**


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Count** method to determine how many validation issues exist in the collection of validation issues in the active document.


```vb
Set vsoDocument = Visio.ActiveDocument 
Set vsoIssues = vsoDocument.Validation.Issues
intIssueTotal = vsoIssues.Count
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]