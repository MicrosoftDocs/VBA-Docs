---
title: Workbook.FullNameURLEncoded property (Excel)
keywords: vbaxl10.chm199203
f1_keywords:
- vbaxl10.chm199203
ms.prod: excel
api_name:
- Excel.Workbook.FullNameURLEncoded
ms.assetid: 589d98f7-e6fa-bc28-2c8f-7cb72009737a
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.FullNameURLEncoded property (Excel)

Returns a **String** indicating the name of the object, including its path on disk, as a string. Read-only.


## Syntax

_expression_.**FullNameURLEncoded**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

In this example, Microsoft Excel displays the path and file name of the active workbook to the user.

```vb
Sub UseCanonical() 
 
 ' Display the full path to user. 
 MsgBox ActiveWorkbook.FullNameURLEncoded 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]