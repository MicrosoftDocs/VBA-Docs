---
title: Printer.Port property (Access)
keywords: vbaac10.chm12865
f1_keywords:
- vbaac10.chm12865
ms.prod: access
api_name:
- Access.Printer.Port
ms.assetid: 0fef85fb-fbe7-eada-1629-d56b6008e039
ms.date: 03/23/2019
localization_priority: Normal
---


# Printer.Port property (Access)

Returns a **String** indicating the port name of the specified printer. Read-only.


## Syntax

_expression_.**Port**

_expression_ A variable that represents a **[Printer](Access.Printer.md)** object.


## Example

The following example displays information about all the printers available to the system.

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