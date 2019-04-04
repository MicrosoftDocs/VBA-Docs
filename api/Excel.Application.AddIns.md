---
title: Application.AddIns property (Excel)
keywords: vbaxl10.chm132081
f1_keywords:
- vbaxl10.chm132081
ms.prod: excel
api_name:
- Excel.Application.AddIns
ms.assetid: 0798690a-910a-b832-e143-df51d7c061ca
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AddIns property (Excel)

Returns an **[AddIns](Excel.AddIns.md)** collection that represents all the add-ins listed in the **Add-Ins** dialog box (**Add-Ins** command on the **Developer** tab). Read-only.


## Syntax

_expression_.**AddIns**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Using this method without an object qualifier is equivalent to Application.Addins.


## Example

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the **AddIns** collection is the title of the add-in, not the add-in's file name.


```vb
If AddIns("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]