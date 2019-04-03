---
title: Application.AddIns2 property (Excel)
keywords: vbaxl10.chm133322
f1_keywords:
- vbaxl10.chm133322
ms.prod: excel
api_name:
- Excel.Application.AddIns2
ms.assetid: 3fd3de81-beae-c5b0-572d-c3f81e251db2
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AddIns2 property (Excel)

Returns an **[AddIns2](Excel.AddIns2.md)** collection that represents all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed. Read-only.


## Syntax

_expression_.**AddIns2**

_expression_ A variable that returns an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the **AddIns2** collection is the title of the add-in, not the add-in's file name.

```vb
If Application.AddIns2("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]