---
title: Application.Printers property (Access)
keywords: vbaac10.chm12596
f1_keywords:
- vbaac10.chm12596
ms.prod: access
api_name:
- Access.Application.Printers
ms.assetid: 71383404-8244-6e9b-9c72-8963e0901901
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.Printers property (Access)

Returns the **[Printers](Access.Printers.md)** collection representing all the available printers on the current system. Read-only **Printers** collection.


## Syntax

_expression_.**Printers**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Example

The following example displays information about all the printers available on the current system.


```vb
Dim prtLoop As Printer 
 
For Each prtLoop In Application.Printers 
 With prtLoop 
 MsgBox "Device name: " & .DeviceName & vbCr _ 
 & "Driver name: " & .DriverName & vbCr _ 
 & "Port: " & .Port 
 End With 
Next prtLoop 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]