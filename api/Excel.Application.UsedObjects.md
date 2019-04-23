---
title: Application.UsedObjects property (Excel)
keywords: vbaxl10.chm133264
f1_keywords:
- vbaxl10.chm133264
ms.prod: excel
api_name:
- Excel.Application.UsedObjects
ms.assetid: bf214478-990b-35c8-1e23-a9d1732e7ef3
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.UsedObjects property (Excel)

Returns a **[UsedObjects](Excel.UsedObjects.md)** object representing objects allocated in a workbook. Read-only.


## Syntax

_expression_.**UsedObjects**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, Microsoft Excel determines the quantity of objects that have been allocated and notifies the user. This example assumes a recalculation was performed in the application and was interrupted before finishing.

```vb
Sub CountUsedObjects() 
 
 MsgBox "The number of used objects in this application is: " & _ 
 Application.UsedObjects.Count 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]