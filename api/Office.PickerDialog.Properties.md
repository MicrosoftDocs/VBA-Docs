---
title: PickerDialog.Properties property (Office)
keywords: vbaof11.chm340003
f1_keywords:
- vbaof11.chm340003
ms.prod: office
api_name:
- Office.PickerDialog.Properties
ms.assetid: 053b5d62-9d9a-68ed-c7ed-cf4df7053ecc
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerDialog.Properties property (Office)

Returns the **[PickerProperties](office.pickerproperties.md)** object to specify custom properties for the data handler component. Read-only.


## Syntax

_expression_.**Properties**

_expression_ An expression that returns a **[PickerDialog](Office.PickerDialog.md)** object.


## Remarks

The properties of the **PickerProperties** object will be passed to the data handler.


## Example

The following code sets various **PickerDialog** properties and retrieves the results.


```vb
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
Set objPickerProperties = objPickerDialog.Properties 
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "https://my", msoPickerFieldtypeText) 
 
' Show the Picker Dialog with no existing result. 
Set objPickerResults = objPickerDialog.Show(True) 

```


## See also

- [PickerDialog object members](overview/Library-Reference/pickerdialog-members-office.md)
- [PickerDialog interface](https://docs.microsoft.com/dotnet/api/microsoft.office.core.pickerdialog?view=office-pia)
- [Object Picker dialog box interfaces](https://docs.microsoft.com/windows/desktop/ad/object-picker-dialog-box-interfaces)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]