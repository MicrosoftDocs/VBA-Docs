---
title: Application.Views method (Project)
keywords: vbapj.chm301
f1_keywords:
- vbapj.chm301
ms.prod: project-server
api_name:
- Project.Application.Views
ms.assetid: 76f29c4c-1854-e136-2d72-d50fe786c26b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Views method (Project)

Displays the **More Views** dialog box with the current view selected, which prompts the user to manage views.


## Syntax

_expression_. `Views`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The **Views** method has the same effect as the **More Views** command in the **Other Views** drop-down list on the **View** tab of the Ribbon.

To specify the pane to select in a split view, use the **[ViewsEx](Project.Application.ViewsEx.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]