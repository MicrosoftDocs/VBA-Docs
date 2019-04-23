---
title: Application.DisplayFullScreen property (Excel)
keywords: vbaxl10.chm133121
f1_keywords:
- vbaxl10.chm133121
ms.prod: excel
api_name:
- Excel.Application.DisplayFullScreen
ms.assetid: b42708ea-a273-c38a-5a61-d15e26c14fed
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DisplayFullScreen property (Excel)

**True** if Microsoft Excel is in full-screen mode. Read/write **Boolean**.


## Syntax

_expression_.**DisplayFullScreen**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Full-screen mode maximizes the application window so that it fills the entire screen and hides the application title bar. Toolbars, the status bar, and the formula bar maintain separate display settings for full-screen mode and normal mode.


## Example

This example sets Microsoft Excel to be displayed in full-screen mode.

```vb
Application.DisplayFullScreen = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
