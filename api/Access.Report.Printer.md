---
title: Report.Printer property (Access)
keywords: vbaac10.chm13810
f1_keywords:
- vbaac10.chm13810
ms.prod: access
api_name:
- Access.Report.Printer
ms.assetid: 9e21b583-5539-bc24-49a0-c248e7f9aafb
ms.date: 03/20/2019
localization_priority: Normal
---


# Report.Printer property (Access)

Returns or sets a **[Printer](Access.Printer.md)** object representing the default printer on the current system. Read/write.


## Syntax

_expression_.**Printer**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


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