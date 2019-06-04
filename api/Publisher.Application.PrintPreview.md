---
title: Application.PrintPreview property (Publisher)
keywords: vbapb10.chm131106
f1_keywords:
- vbapb10.chm131106
ms.prod: publisher
api_name:
- Publisher.Application.PrintPreview
ms.assetid: a6606819-89d1-609d-62c3-c59159ff2ef7
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.PrintPreview property (Publisher)

**True** to display in Print Preview the publication in the current view. Read/write **Boolean**.


## Syntax

_expression_.**PrintPreview**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

Boolean


## Example

This example switches the view to Print Preview.

```vb
Sub GoToPrintPreview() 
 PrintPreview = True 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]