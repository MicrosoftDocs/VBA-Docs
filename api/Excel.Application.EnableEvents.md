---
title: Application.EnableEvents property (Excel)
keywords: vbaxl10.chm133240
f1_keywords:
- vbaxl10.chm133240
ms.prod: excel
api_name:
- Excel.Application.EnableEvents
ms.assetid: 5e14ce7b-02f6-03d4-2dfc-1df05a032301
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EnableEvents property (Excel)

 **True** if events are enabled for the specified object. Read/write **Boolean**.


## Syntax

_expression_. `EnableEvents`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example disables events before a file is saved so that the  **BeforeSave** event doesn't occur.


```vb
Application.EnableEvents = False 
ActiveWorkbook.Save 
Application.EnableEvents = True
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
