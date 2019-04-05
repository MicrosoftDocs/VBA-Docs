---
title: Application.ProtectedViewWindowDeactivate event (Excel)
keywords: vbaxl10.chm504113
f1_keywords:
- vbaxl10.chm504113
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowDeactivate
ms.assetid: 39df50ca-53e0-784a-a803-e9ac6f456d11
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ProtectedViewWindowDeactivate event (Excel)

Occurs when a Protected View window is deactivated.


## Syntax

_expression_.**ProtectedViewWindowDeactivate** (_Pvw_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)**|An object that represents the deactivated Protected View window.|

## Return value

Nothing


## Example

The following code example minimizes any Protected View window when it is deactivated. This code must be placed in a class module, and an instance of that class must be correctly initialized. 

For more information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


```vb
Private Sub App_ProtectedViewWindowDeactivate(ByVal Pvw As ProtectedViewWindow) 
 Pvw.WindowState = xlProtectedViewWindowMinimized 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]