---
title: PickerDialog.Title property (Office)
keywords: vbaof11.chm340002
f1_keywords:
- vbaof11.chm340002
ms.prod: office
api_name:
- Office.PickerDialog.Title
ms.assetid: 76531e47-91a4-4d82-7825-ab900c5bf8e2
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerDialog.Title property (Office)

Sets or returns the title of a **PickerDialog** displayed in the **PickerDialog**. Read/write.


## Syntax

_expression_.**Title**

_expression_ An expression that returns a **[PickerDialog](Office.PickerDialog.md)** object.


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]