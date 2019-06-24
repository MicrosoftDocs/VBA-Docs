---
title: InvisibleApp.AlertResponse property (Visio)
keywords: vis_sdr.chm17513070
f1_keywords:
- vis_sdr.chm17513070
ms.prod: visio
api_name:
- Visio.InvisibleApp.AlertResponse
ms.assetid: eb0fabb1-809e-b952-da99-e18eda0f6970
ms.date: 06/24/2019
localization_priority: Normal
---


# InvisibleApp.AlertResponse property (Visio)

Determines whether Microsoft Visio shows alerts and modal dialog boxes to the user. Read/write.


## Syntax

_expression_.**AlertResponse**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Integer


## Remarks

Certain operations, such as closing a document that has unsaved modifications, cause Visio to display an alert or modal dialog box requesting the user to supply a response such as **OK**, **Yes**, **No**, or **Cancel**. To prevent Visio from displaying an alert or modal dialog box when a program performs such an action, set the **AlertResponse** property to a default value for the response. In this case, Visio does not display the alert or modal dialog box; instead, Visio behaves as if the user responded to the alert or modal dialog box with the value of the **AlertResponse** property.

If the **AlertResponse** property is 0 (its default value), alerts and modal dialog boxes are displayed.

The values that you supply for the **AlertResponse** property correspond to the standard Windows constants IDOK, IDCANCEL, and so forth.

|Constant|Value|
|:-----|:-----|
|IDOK|1|
|IDCANCEL|2|
|IDABORT|3|
|IDRETRY|4|
|IDIGNORE|5|
|IDYES|6|
|IDNO|7|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]