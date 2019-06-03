---
title: AdvancedPrintOptions.PrintCropMarks property (Publisher)
keywords: vbapb10.chm7077895
f1_keywords:
- vbapb10.chm7077895
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.PrintCropMarks
ms.assetid: 0b777632-572c-7080-8f4d-b97a284d04e2
ms.date: 06/04/2019
localization_priority: Normal
---


# AdvancedPrintOptions.PrintCropMarks property (Publisher)

**True** to print crop marks for the specified publication. The default is **True**. Read/write **Boolean**.


## Syntax

_expression_.**PrintCropMarks**

_expression_ A variable that represents an **[AdvancedPrintOptions](Publisher.AdvancedPrintOptions.md)** object.


## Return value

Boolean


## Remarks

This property corresponds to the **Crop marks** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.

Crop marks are used as guides when a printed publication is trimmed to its intended size.

These printer's marks print outside the publication and can only be printed if the size of the sheet being printed to is larger than the publication page size.


## Example

The following example sets crop marks and job information to print with the publication. If the publication is printed as separations, the additional types of printer's marks are also set to print. This example assumes that the size of the paper being printed to is larger than the publication page size.


```vb
Sub SetPrintersMarksToPrint() 
 With ActiveDocument.AdvancedPrintOptions 
 .PrintCropMarks = True 
 .PrintJobInformation = True 
 If PrintMode = pbPrintModeSeparations Then 
 .PrintRegistrationMarks = True 
 .PrintDensityBars = True 
 .PrintColorBars = True 
 End If 
 End With 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]