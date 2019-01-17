---
title: Application.FindFile method (Excel)
keywords: vbaxl10.chm133256
f1_keywords:
- vbaxl10.chm133256
ms.prod: excel
api_name:
- Excel.Application.FindFile
ms.assetid: c4367047-0f7d-1bda-5103-f26eedd98e8a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FindFile method (Excel)

Displays the  **Open** dialog box.


## Syntax

_expression_. `FindFile`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Return value

Boolean


## Remarks

This method displays the  **Open** dialog box and allows the user to open a file. If a new file is opened successfully, this method returns **True**. If the user cancels the dialog box, this method returns **False**.


## Example

This example displays the  **Open** dialog box.


```vb
Application.FindFile
```


## See also


[Application Object](Excel.Application(object).md)

