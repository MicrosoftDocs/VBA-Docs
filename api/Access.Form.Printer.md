---
title: Form.Printer property (Access)
keywords: vbaac10.chm13523
f1_keywords:
- vbaac10.chm13523
ms.prod: access
api_name:
- Access.Form.Printer
ms.assetid: c533271a-c500-57de-f16c-ed384698f829
ms.date: 03/14/2019
localization_priority: Normal
---


# Form.Printer property (Access)

Returns or sets a **[Printer](Access.Printer.md)** object representing the default printer on the current system. Read/write.


## Syntax

_expression_.**Printer**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Example

The following example makes the first printer in the **[Printers](Access.Printers.md)** collection the default printer for the system, and then reports its name, driver information, and port information.

```vb
Dim prtDefault As Printer 
 
Set Application.Printer = Application.Printers(0) 
 
Set prtDefault = Application.Printer 
 
With prtDefault 
 MsgBox "Device name: " & .DeviceName & vbCr _ 
 & "Driver name: " & .DriverName & vbCr _ 
 & "Port: " & .Port 
End With 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]