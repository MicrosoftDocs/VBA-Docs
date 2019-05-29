---
title: Workbook.BeforeSave event (Excel)
keywords: vbaxl10.chm503077
f1_keywords:
- vbaxl10.chm503077
ms.prod: excel
api_name:
- Excel.Workbook.BeforeSave
ms.assetid: dfa3e20f-1fb2-f84f-4b92-a98f22b6e637
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.BeforeSave event (Excel)

Occurs before the workbook is saved.


## Syntax

_expression_.**BeforeSave** (_SaveAsUI_, _Cancel_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveAsUI_|Required| **Boolean**| **True** if the **Save As** dialog box is displayed due to changes made that need to be saved in the workbook.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the workbook isn't saved when the procedure is finished.|

## Return value

**Nothing**


## Example

This example prompts the user for a yes or no response before saving the workbook.

```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, _ 
        Cancel as Boolean) 
    a = MsgBox("Do you really want to save the workbook?", vbYesNo) 
    If a = vbNo Then Cancel = True 
End Sub
```

<br/>

This example uses the **BeforeSave** event to verify that certain cells contain data before the workbook can be saved. The workbook cannot be saved until there is data in each of the following cells: D5, D7, D9, D11, D13, and D15.

```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
   'If the six specified cells do not contain data, then display a message box with an error
   'and cancel the attempt to save.
   If WorksheetFunction.CountA(Worksheets("Sheet1").Range("D5,D7,D9,D11,D13,D15")) < 6 Then
      MsgBox "Workbook will not be saved unless" & vbCrLf & _
      "All required fields have been filled in!"
      Cancel = True
   End If
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
