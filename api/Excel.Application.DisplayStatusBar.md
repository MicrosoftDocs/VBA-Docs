---
title: Application.DisplayStatusBar property (Excel)
keywords: vbaxl10.chm133127
f1_keywords:
- vbaxl10.chm133127
ms.prod: excel
api_name:
- Excel.Application.DisplayStatusBar
ms.assetid: bf70a679-bd50-cce7-0dc0-0dc57835038c
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DisplayStatusBar property (Excel)

**True** if the status bar is displayed. Read/write **Boolean**.


## Syntax

_expression_.**DisplayStatusBar**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example saves the current state of the **DisplayStatusBar** property, and then sets the property to **True** so that the status bar is visible.

```vb
saveStatusBar = Application.DisplayStatusBar 
Application.DisplayStatusBar = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]