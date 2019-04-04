---
title: Application.EnableEvents property (Excel)
keywords: vbaxl10.chm133240
f1_keywords:
- vbaxl10.chm133240
ms.prod: excel
api_name:
- Excel.Application.EnableEvents
ms.assetid: 5e14ce7b-02f6-03d4-2dfc-1df05a032301
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.EnableEvents property (Excel)

**True** if events are enabled for the specified object. Read/write **Boolean**.


## Syntax

_expression_.**EnableEvents**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example disables events before a file is saved so that the **BeforeSave** event doesn't occur.


```vb
Application.EnableEvents = False 
ActiveWorkbook.Save 
Application.EnableEvents = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
