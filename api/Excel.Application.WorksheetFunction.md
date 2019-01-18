---
title: Application.WorksheetFunction property (Excel)
keywords: vbaxl10.chm183116
f1_keywords:
- vbaxl10.chm183116
ms.prod: excel
api_name:
- Excel.Application.WorksheetFunction
ms.assetid: fd1333bf-8739-303d-30b4-85a99fb344b3
ms.date: 06/08/2017
localization_priority: Priority
---


# Application.WorksheetFunction property (Excel)

Returns the  **[WorksheetFunction](Excel.WorksheetFunction.md)** object. Read-only.


## Syntax

_expression_. `WorksheetFunction`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example displays the result of applying the  **Min** worksheet function to the range A1:A10.


```vb
Set myRange = Worksheets("Sheet1").Range("A1:C10") 
answer = Application.WorksheetFunction.Min(myRange) 
MsgBox answer
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]