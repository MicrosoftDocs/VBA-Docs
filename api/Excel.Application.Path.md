---
title: Application.Path property (Excel)
keywords: vbaxl10.chm133189
f1_keywords:
- vbaxl10.chm133189
api_name:
- Excel.Application.Path
ms.assetid: 0ef5d0fc-f46a-c133-232a-8a20cf2d4034
ms.date: 04/05/2019
ms.localizationpriority: medium
---


# Application.Path property (Excel)

Returns a **String** value that represents the complete path to the application, excluding the final separator and name of the application.


## Syntax

_expression_.**Path**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the complete path to Microsoft Excel.

```vb
Sub TotalPath() 
 
 MsgBox "The path is " & Application.Path 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
