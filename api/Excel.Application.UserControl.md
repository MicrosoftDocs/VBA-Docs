---
title: Application.UserControl property (Excel)
keywords: vbaxl10.chm133224
f1_keywords:
- vbaxl10.chm133224
ms.prod: excel
api_name:
- Excel.Application.UserControl
ms.assetid: fd55727d-8f79-14bf-038b-a31a56829a55
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.UserControl property (Excel)

**True** if the application is visible or if it was created or started by the user. **False** if you created or started the application programmatically by using the **CreateObject** or **GetObject** functions, and the application is hidden. Read/write **Boolean**.


## Syntax

_expression_.**UserControl**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

When the **UserControl** property is **False** for an object, that object is released when the last programmatic reference to the object is released. If this property is **False**, Microsoft Excel quits when the last object in the session is released.


## Example

This example displays the status of the **UserControl** property.

```vb
If Application.UserControl Then 
 MsgBox "This workbook was created by the user" 
Else 
 MsgBox "This workbook was created programmatically" 
End If 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]