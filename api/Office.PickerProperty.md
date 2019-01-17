---
title: PickerProperty object (Office)
keywords: vbaof11.chm336000
f1_keywords:
- vbaof11.chm336000
ms.prod: office
api_name:
- Office.PickerProperty
ms.assetid: fd3702fe-bf03-f22c-78c2-ac6c47a1d028
ms.date: 06/08/2017
localization_priority: Normal
---


# PickerProperty object (Office)

Represents an object for passing a custom property. p


## Example

The following code sets the Picker Dialog properties and then displays the Picker dialog.


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


## Properties



|Name|
|:-----|
|[Application](Office.PickerProperty.Application.md)|
|[Creator](Office.PickerProperty.Creator.md)|
|[Id](Office.PickerProperty.Id.md)|
|[Type](Office.PickerProperty.Type.md)|
|[Value](Office.PickerProperty.Value.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
