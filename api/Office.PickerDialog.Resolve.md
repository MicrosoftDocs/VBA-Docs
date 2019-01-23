---
title: PickerDialog.Resolve method (Office)
keywords: vbaof11.chm340006
f1_keywords:
- vbaof11.chm340006
ms.prod: office
api_name:
- Office.PickerDialog.Resolve
ms.assetid: 50b1792a-ecf0-ab66-6a9d-7f72c788d859
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerDialog.Resolve method (Office)

Resolves the token by using the **PickerDialog** and retrieves the results.


## Syntax

_expression_.**Resolve** (_TokenText_, _duplicateDlgMode_)

_expression_ An expression that returns a **[PickerDialog](Office.PickerDialog.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TokenText_|Required|**String**|The text string to resolve.|
| _duplicateDlgMode_|Required|**Integer**||

## Return value

PickerResults


## Example

Resolves entities by using the **PickerDialog** object.


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
 
' Resolve the token by using Picker Dialog and get the results. 
Set objPickerResults = objPickerDialog.Resolve("johndoe", False) 

```

## See also

- [PickerDialog object members](overview/Library-Reference/pickerdialog-members-office.md)
- [PickerDialog interface](https://docs.microsoft.com/dotnet/api/microsoft.office.core.pickerdialog?view=office-pia)
- [Object Picker dialog box interfaces](https://docs.microsoft.com/windows/desktop/ad/object-picker-dialog-box-interfaces)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]