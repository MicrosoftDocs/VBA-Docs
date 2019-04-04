---
title: Application.GenerateGetPivotData property (Excel)
keywords: vbaxl10.chm133275
f1_keywords:
- vbaxl10.chm133275
ms.prod: excel
api_name:
- Excel.Application.GenerateGetPivotData
ms.assetid: 83effd5f-5101-ba1b-ab45-722e26074ea7
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.GenerateGetPivotData property (Excel)

Returns **True** when Microsoft Excel can get PivotTable report data. Read/write **Boolean**.


## Syntax

_expression_.**GenerateGetPivotData**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In the following example, Microsoft Excel determines the status of getting PivotTable report data and notifies the user. This example assumes a PivotTable report exists on the active worksheet.


```vb
Sub PivotTableInfo() 
 
 ' Determine the ability to get PivotTable report data and notify user. 
 If Application.GenerateGetPivotData = True Then 
 MsgBox "The ability to get PivotTable report data is enabled." 
 Else 
 Msgbox "The ability to get PivotTable report data is disabled." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]