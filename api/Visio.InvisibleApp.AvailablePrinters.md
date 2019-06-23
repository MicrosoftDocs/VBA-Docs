---
title: InvisibleApp.AvailablePrinters property (Visio)
keywords: vis_sdr.chm17550510
f1_keywords:
- vis_sdr.chm17550510
ms.prod: visio
api_name:
- Visio.InvisibleApp.AvailablePrinters
ms.assetid: 3e4bba9c-d338-deea-ef78-7150804d0216
ms.date: 06/24/2019
localization_priority: Normal
---


# InvisibleApp.AvailablePrinters property (Visio)

Returns a list of installed printers. Read-only.


## Syntax

_expression_.**AvailablePrinters**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

String()


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **AvailablePrinters** property to get a list of available printers.

```vb
Public Sub AvailablePrinters_example() 
 
 Dim aStrPrinters() As String 
 Dim strPrinter As Variant 
 
 aStrPrinters = Application.AvailablePrinters 
 
 For Each strPrinter In aStrPrinters 
 
 Debug.Print strPrinter 
 
 Next 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]