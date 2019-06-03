---
title: Printers object (Access)
keywords: vbaac10.chm12881
f1_keywords:
- vbaac10.chm12881
ms.prod: access
api_name:
- Access.Printers
ms.assetid: 5200c507-75ae-f9a8-c737-c28e175e7ea4
ms.date: 03/21/2019
localization_priority: Normal
---


# Printers object (Access)

The **Printers** collection contains **[Printer](Access.Printer.md)** objects representing all the printers available on the current system.


## Remarks

Use the **[Printers](Access.Application.Printers.md)** property of the **Application** object to return the **Printers** collection. You can enumerate through the **Printers** collection by using the **For Each...Next** statement.

You can refer to an individual **Printer** object in the **Printers** collection either by referring to the printer by name, or by referring to its index within the collection.

The **Printers** collection is indexed beginning with zero. If you refer to a printer by its index, the first printer is Printers(0), the second printer is Printers(1), and so on.

You can't add or delete a **Printer** object from the **Printers** collection.


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


## Properties

- [Application](Access.Printers.Application.md)
- [Count](Access.Printers.Count.md)
- [Item](Access.Printers.Item.md)
- [Parent](Access.Printers.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]