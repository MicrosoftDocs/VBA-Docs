---
title: PickerDialog.CreatePickerResults method (Office)
keywords: vbaof11.chm340004
f1_keywords:
- vbaof11.chm340004
ms.prod: office
api_name:
- Office.PickerDialog.CreatePickerResults
ms.assetid: 39954f3e-53ef-f33c-9e90-a2247fd7882a
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerDialog.CreatePickerResults method (Office)

Creates an empty **[PickerResults](office.pickerresults.md)** object.


## Syntax

_expression_.**CreatePickerResults**

_expression_ An expression that returns a **[PickerDialog](Office.PickerDialog.md)** object.


## Return value

PickerResults


## Remarks

You can add the **PickerResult** to the returned object and specify it to the second parameter of the **Show** method as already existing results of the **PickerDialog** object.


## Example

The following code sets various properties of the **PickerDialog** and adds the already existing **PickerResults** to the results.


```vb
Dim objPickerDialog As PickerDialog 
Dim objPickerExistingResults As PickerResults 
 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
 
Set objPickerExistingResults = objPickerDialog.CreatePickerResults 
Set objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User") 
Set objPickerResults = objPickerDialog.Show(True, objPickerExistingResult) 

```


## See also

- [PickerDialog object members](overview/Library-Reference/pickerdialog-members-office.md)
- [PickerDialog interface](https://docs.microsoft.com/dotnet/api/microsoft.office.core.pickerdialog?view=office-pia)
- [Object Picker dialog box interfaces](https://docs.microsoft.com/windows/desktop/ad/object-picker-dialog-box-interfaces)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]