---
title: Worksheet.Index property (Excel)
keywords: vbaxl10.chm174078
f1_keywords:
- vbaxl10.chm174078
api_name:
- Excel.Worksheet.Index
ms.assetid: 970065b3-f9bd-d518-261a-f5f704c350df
ms.date: 05/30/2019
ms.localizationpriority: medium
---


# Worksheet.Index property (Excel)

Returns a **Long** value that represents the index number of the object within the collection of similar objects.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example displays the tab number of the sheet specified by the name that you type. For example, if Sheet4 is the third tab in the active workbook, the example displays "3" in a message box.

```vb
Sub DisplayTabNumber() 
 Dim strSheetName as String 
 
 strSheetName = InputBox("Type a sheet name, such as Sheet4.") 
 
 MsgBox "This sheet is tab number " & Sheets(strSheetName).Index 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
