---
title: AddIn.FullName property (Excel)
keywords: vbaxl10.chm185075
f1_keywords:
- vbaxl10.chm185075
ms.prod: excel
api_name:
- Excel.AddIn.FullName
ms.assetid: d5e0672e-0595-16f7-9364-f8aee9d9388e
ms.date: 04/03/2019
localization_priority: Normal
---


# AddIn.FullName property (Excel)

Returns the name of the object, including its path on disk, as a string. Read-only **String**.


## Syntax

_expression_.**FullName**

_expression_ A variable that represents an **[AddIn](Excel.AddIn.md)** object.


## Example

This example displays the path and file name of every available add-in.

```vb
For Each a In AddIns 
 MsgBox a.FullName 
Next a
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]