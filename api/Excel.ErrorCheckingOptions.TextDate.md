---
title: ErrorCheckingOptions.TextDate property (Excel)
keywords: vbaxl10.chm698076
f1_keywords:
- vbaxl10.chm698076
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.TextDate
ms.assetid: eb251a44-4dac-01e5-1d01-b4e8bd71e8e2
ms.date: 04/26/2019
localization_priority: Normal
---


# ErrorCheckingOptions.TextDate property (Excel)

When set to **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, cells that contain a text date with a two-digit year. **False** disables error checking for cells containing a text date with a two-digit year. Read/write **Boolean**.


## Syntax

_expression_.**TextDate**

_expression_ A variable that represents an **[ErrorCheckingOptions](Excel.ErrorCheckingOptions.md)** object.


## Example

In the following example, the **AutoCorrect Options** button appears for cell A1, which contains a text date with a two-digit year.

```vb
Sub CheckTextDate() 
 
 ' Simulate an error by referencing a text date with a two-digit year. 
 Application.ErrorCheckingOptions.TextDate = True 
 Range("A1").Formula = "'April 23, 00" 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]