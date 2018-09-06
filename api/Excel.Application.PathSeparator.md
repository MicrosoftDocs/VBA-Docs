---
title: Application.PathSeparator Property (Excel)
keywords: vbaxl10.chm133190
f1_keywords:
- vbaxl10.chm133190
ms.prod: excel
api_name:
- Excel.Application.PathSeparator
ms.assetid: 573ef52d-3f1a-4fa3-4d4b-f047be67f26f
ms.date: 06/08/2017
---


# Application.PathSeparator Property (Excel)

Returns the path separator character ("\\"). Read-only  **String** .


## Syntax

 _expression_. `PathSeparator`

 _expression_ A variable that represents an [Application](Excel.Application(Graph property).md) object.


## Example

This example displays the current path separator.


```vb
MsgBox "The path separator character is " & _ 
 Application.PathSeparator
```


## See also


[Application Object](Excel.Application(object).md)

