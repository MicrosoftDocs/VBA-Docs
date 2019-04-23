---
title: Application.ShowQuickAnalysis property (Excel)
keywords: vbaxl10.chm133337
f1_keywords:
- vbaxl10.chm133337
ms.prod: excel
ms.assetid: 043d9523-1fbc-0afd-2adf-9775e71058c0
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ShowQuickAnalysis property (Excel)

Controls whether the Quick Analysis contextual user interface is displayed on selection. **True** means that the **Quick Analysis** button will show. 

Corresponds to the **Show Quick Analysis options on selection** check box located on the **File** menu > **Options** > **Excel Options** > **General** tab. Read/write **Boolean**.


## Syntax

_expression_.**ShowQuickAnalysis**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Property value

**BOOL**


## Example

This example hides the **Quick Analysis** button on selection.

```vb
Application.ShowQuickAnalysis = False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]