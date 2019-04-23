---
title: Application.DisplayNoteIndicator property (Excel)
keywords: vbaxl10.chm133122
f1_keywords:
- vbaxl10.chm133122
ms.prod: excel
api_name:
- Excel.Application.DisplayNoteIndicator
ms.assetid: 96d43af3-0ceb-4bc2-ebaf-33cbe3e30a8a
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DisplayNoteIndicator property (Excel)

**True** if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners). Read/write **Boolean**.


## Syntax

_expression_.**DisplayNoteIndicator**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example hides note indicators.


```vb
Application.DisplayNoteIndicator = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]