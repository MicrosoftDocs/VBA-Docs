---
title: Application.VisualReportsAdditionalTemplatePath property (Project)
keywords: vbapj.chm131391
f1_keywords:
- vbapj.chm131391
ms.prod: project-server
api_name:
- Project.Application.VisualReportsAdditionalTemplatePath
ms.assetid: d1727b8c-595e-bf41-cbd5-3cebed893636
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.VisualReportsAdditionalTemplatePath property (Project)

Gets or sets the additional path for Visual Reports templates. Read/write  **String**.


## Syntax

_expression_. `VisualReportsAdditionalTemplatePath`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

The **Include report templates from** text box in the **Visual Reports - Create Report** dialog box shows the value of the **VisualReportsAdditionalTemplatePath** property.

To clear the additional path and template name, use an empty string ("").


> [!NOTE] 
> When you set a path value with the **VisualReportsAdditionalTemplatePath** property, Project does not check whether the path exists.


## Example

The following example sets the additional path to "C:\My Templates".


```vb
Application.VisualReportsAdditionalTemplatePath = "C:\My Templates"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]