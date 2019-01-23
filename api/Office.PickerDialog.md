---
title: PickerDialog object (Office)
keywords: vbaof11.chm340000
f1_keywords:
- vbaof11.chm340000
ms.prod: office
api_name:
- Office.PickerDialog
ms.assetid: 279b1a6a-f09d-a0e7-89c9-aac6c581439f
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerDialog object (Office)

Provides dialog user interface functionality for picking people or picking data.

## Remarks

Get the **PickerDialog** object through the **PickerDialog** property in the **Application** object.


## Example

The following code sets the **PickerDialog** properties and then displays the **PickerDialog**.


```vb
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
Dim objPickerProperty As PickerProperty 
Dim objPickerExistingResults As PickerResults 
Dim objPickerExistingResult As PickerResult 
Dim objPickerResults As PickerResults 
 
' Configure the Picker Dialog properties. 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
Set objPickerProperties = objPickerDialog.Properties 
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "https://my", msoPickerFieldtypeText) 
Set objPickerExistingResults = objPickerDialog.CreatePickerResults 
Set objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User") 
 
' Show the Picker Dialog and get the results. 
Set objPickerResults = objPickerDialog.Show(True, objPickerExistingResult)
```



## See also

- [PickerDialog object members](overview/Library-Reference/pickerdialog-members-office.md)
- [PickerDialog interface](https://docs.microsoft.com/dotnet/api/microsoft.office.core.pickerdialog?view=office-pia)
- [Object Picker dialog box interfaces](https://docs.microsoft.com/windows/desktop/ad/object-picker-dialog-box-interfaces)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]