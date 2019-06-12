---
title: Printer object (Publisher)
keywords: vbapb10.chm9043967
f1_keywords:
- vbapb10.chm9043967
ms.prod: publisher
api_name:
- Publisher.Printer
ms.assetid: 46f8c6a2-4cf1-bb6a-1214-a751440870f2
ms.date: 06/01/2019
localization_priority: Normal
---


# Printer object (Publisher)

A **Printer** object represents a printer installed on your computer.


## Remarks

Many of the properties of the **Printer** object, such as **PaperSize**, **PaperSource**, and **PaperOrientation**, correspond to the settings in the **Print Setup** dialog box (**File** menu) in the Microsoft Publisher user interface.

The collection of all the printers installed on your computer is represented by the **InstalledPrinters** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you can use the **PrinterName** and **IsActivePrinter** properties of the **Printer** object to get a list of all the installed printers on the computer, determine which of them is currently the active printer, and get some of the settings of the active printer. The macro displays the results in the Immediate window.

```vb
Public Sub Printer_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 Debug.Print pubPrinter.PrinterName 
 If pubPrinter.IsActivePrinter Then 
 Debug.Print "This is the active printer" 
 Debug.Print "Paper size is ", pubPrinter.PaperSize 
 Debug.Print "Paper orientation is ", pubPrinter.PaperOrientation 
 Debug.Print "Paper source is ", pubPrinter.PaperSource 
 End If 
 Next 
 
End Sub
```


## Properties

- [Application](Publisher.Printer.Application.md)
- [DriverType](Publisher.Printer.DriverType.md)
- [Index](Publisher.Printer.Index.md)
- [IsActivePrinter](Publisher.Printer.IsActivePrinter.md)
- [IsColor](Publisher.Printer.IsColor.md)
- [IsDuplex](Publisher.Printer.IsDuplex.md)
- [PaperHeight](Publisher.Printer.PaperHeight.md)
- [PaperOrientation](Publisher.Printer.PaperOrientation.md)
- [PaperSize](Publisher.Printer.PaperSize.md)
- [PaperSource](Publisher.Printer.PaperSource.md)
- [PaperWidth](Publisher.Printer.PaperWidth.md)
- [Parent](Publisher.Printer.Parent.md)
- [PrintableRect](Publisher.Printer.PrintableRect.md)
- [PrinterName](Publisher.Printer.PrinterName.md)
- [PrintMode](Publisher.Printer.PrintMode.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]