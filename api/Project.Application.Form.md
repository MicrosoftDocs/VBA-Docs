---
title: Application.Form method (Project)
keywords: vbapj.chm1004
f1_keywords:
- vbapj.chm1004
ms.prod: project-server
api_name:
- Project.Application.Form
ms.assetid: 23e7c800-bda9-c931-bc27-084dec872953
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Form method (Project)

Displays a custom form. The **Form** method produces an error if a resource form is specified when the active view is a task view, and vice versa.


## Syntax

_expression_. `Form`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a custom form. The default is a task form when the active view is a task view, and a resource form when the active view is a resource view.|

## Return value

 **Boolean**


## Example

The following example displays the Cost Tracking form.


```vb
Sub DisplayCostTrackingForm 
 Form("Cost Tracking") 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]