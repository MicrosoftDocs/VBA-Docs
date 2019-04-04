---
title: Application.ProtectedViewWindowOpen event (Excel)
keywords: vbaxl10.chm504108
f1_keywords:
- vbaxl10.chm504108
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowOpen
ms.assetid: 17c847d9-a9d2-28da-832a-01d7719f1248
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ProtectedViewWindowOpen event (Excel)

Occurs when a workbook is opened in a Protected View window.


## Syntax

_expression_.**ProtectedViewWindowOpen** (_Pvw_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)**|An object that represents the Protected View window that is opened.|

## Return value

Nothing


## Example

The following code example informs the user that the workbook will be opened in a Protected View window. This code must be placed in a class module, and an instance of that class must be correctly initialized. 

For more information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


```vb
Private Sub App_ProtectedViewWindowOpen(ByVal Pvw As ProtectedViewWindow) 
 MsgBox "You are opening the following workbook in Protected View: " _ 
 & Pvw.Caption 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]