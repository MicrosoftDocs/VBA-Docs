---
title: Hyperlinks.Delete method (Excel)
keywords: vbaxl10.chm534078
f1_keywords:
- vbaxl10.chm534078
ms.prod: excel
api_name:
- Excel.Hyperlinks.Delete
ms.assetid: 6875e532-a1af-2080-f80e-89d651294db0
ms.date: 06/21/2019
localization_priority: Normal
---


# Hyperlinks.Delete method (Excel)

Deletes the object.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[Hyperlinks](Excel.Hyperlinks.md)** object.


## Remarks

Calling the **Delete** method on the specified **Hyperlinks** object is equivalent to using both the **Clear Hyperlinks** and **Clear Formats** commands from the **Clear** drop-down list in the **Editing** section of the **Home** tab. Not only hyperlinks will be removed; cell formatting will be removed also. If you only want to remove the hyperlink, see the **[Range.ClearHyperlinks](Excel.Range.ClearHyperlinks.md)** method.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
