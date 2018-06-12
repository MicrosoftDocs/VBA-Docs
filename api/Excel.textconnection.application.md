---
title: TextConnection.Application Property (Excel)
keywords: vbaxl10.chm925073
f1_keywords:
- vbaxl10.chm925073
ms.prod: excel
ms.assetid: a3dc9071-4d42-6293-b9df-25dcc84d4ca8
ms.date: 06/08/2017
---


# TextConnection.Application Property (Excel)

Returns an  **[Application](Excel.Application(objec).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a[TextConnection Object (Excel)](Excel.textconnection.md) object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Property value

 **APPLICATION**


## See also


#### Other resources



[TextConnection Object](Excel.textconnection.md)

