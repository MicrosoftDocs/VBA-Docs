---
title: PickerDialog.Show method (Office)
keywords: vbaof11.chm340005
f1_keywords:
- vbaof11.chm340005
ms.prod: office
api_name:
- Office.PickerDialog.Show
ms.assetid: 3073defe-4585-816d-6b86-9959cce4655f
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerDialog.Show method (Office)

Displays the **PickerDialog** with the already specified data handler and given options.


## Syntax

_expression_.**Show** (_IsMultiSelect_, _ExistingResults_)

_expression_ An expression that returns a **[PickerDialog](Office.PickerDialog.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _IsMultiSelect_|Optional|**Boolean**|Specifies whether the **PickerDialog** user interface provides multiple item selection functions.|
| _ExistingResults_|Optional|**PickerResults**|Contains existing **PickerResults** in the **PickerDialog** user interface. These results are displayed in the selected item control.|

## Return value

PickerResults


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