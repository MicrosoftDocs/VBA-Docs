---
title: Printer.DriverName property (Access)
keywords: vbaac10.chm12860
f1_keywords:
- vbaac10.chm12860
ms.prod: access
api_name:
- Access.Printer.DriverName
ms.assetid: 7434f44a-8b55-1f21-e595-363327199037
ms.date: 03/23/2019
localization_priority: Normal
---


# Printer.DriverName property (Access)

Returns a **String** indicating the name of the driver used by the specified printer. Read-only.


## Syntax

_expression_.**DriverName**

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