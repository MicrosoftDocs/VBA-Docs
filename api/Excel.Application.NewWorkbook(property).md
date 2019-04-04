---
title: Application.NewWorkbook property (Excel)
keywords: vbaxl10.chm133283
f1_keywords:
- vbaxl10.chm133283
ms.prod: excel
api_name:
- Excel.Application.NewWorkbook
ms.assetid: 3a50a338-53be-3ac9-d398-c58084e19e6d
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.NewWorkbook property (Excel)

Returns a **[NewFile](Office.NewFile.md)** object.


## Syntax

_expression_.**NewWorkbook**

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


## Return value

NewFile


## Example

In this example, Microsoft Excel sets the variable _wkbOne_ to a **NewFile** object.

```vb
 
Sub SetStartWorking() 
 
    Dim wkbOne As NewFile 
 
    ' Create a reference to an instance of the NewFile object. 
    Set wkbOne = Application.NewWorkbook 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]