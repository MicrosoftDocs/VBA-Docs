---
title: AutoRecover.Path property (Excel)
keywords: vbaxl10.chm696075
f1_keywords:
- vbaxl10.chm696075
ms.prod: excel
api_name:
- Excel.AutoRecover.Path
ms.assetid: 1b95e149-d758-89f9-3879-760ffda01bf8
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoRecover.Path property (Excel)

Returns or sets a  **String** value that represents the complete path to where Microsoft Excel will store the AutoRecover temporary files.


## Syntax

_expression_. `Path`

_expression_ A variable that represents an [AutoRecover](Excel.AutoRecover.md) object.


## Example

This example sets the path of the AutoRecover file to drive C.


```vb
Sub SetPath() 
 
 Application.AutoRecover.Path = "C:\" 
 
End Sub
```


## See also


[AutoRecover Object](Excel.AutoRecover.md)

