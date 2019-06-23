---
title: InvisibleApp.Quit method (Visio)
keywords: vis_sdr.chm17516460
f1_keywords:
- vis_sdr.chm17516460
ms.prod: visio
api_name:
- Visio.InvisibleApp.Quit
ms.assetid: e45406cc-45fb-54a0-6a63-0be0f0647a11
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.Quit method (Visio)

Closes the indicated instance of Microsoft Visio.


## Syntax

_expression_.**Quit**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

**Nothing**


## Remarks

If the  **Quit** method is invoked when any open document has unsaved changes, a dialog box appears asking if you want to save the document. To quit the application without saving and seeing the dialog box, set the **Saved** property of the **Document** object representing the document to **True** immediately before quitting. Set the **Saved** property to **True** only if you are sure you want to close the document without saving changes, because you will lose any unsaved changes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]