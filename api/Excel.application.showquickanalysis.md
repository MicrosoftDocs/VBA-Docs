---
title: Application.ShowQuickAnalysis property (Excel)
keywords: vbaxl10.chm133337
f1_keywords:
- vbaxl10.chm133337
ms.prod: excel
ms.assetid: 043d9523-1fbc-0afd-2adf-9775e71058c0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShowQuickAnalysis property (Excel)

Controls whether the Quick Analysis contextual user interface is displayed on selection.  **TRUE** means the Quick Analysis button will show. Corresponds to the **Show Quick Analysis options on selection** checkbox located in the **File** menu, **Options**,  **Excel Options**, and then  **General** tab. Read/Write. **Boolean**.


## Syntax

_expression_. `ShowQuickAnalysis`

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example hides the Quick Analysis button on selection.


```vb
Application.ShowQuickAnalysis = False
```


## Property value

 **BOOL**


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]