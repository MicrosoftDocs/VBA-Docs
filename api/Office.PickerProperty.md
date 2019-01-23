---
title: PickerProperty object (Office)
keywords: vbaof11.chm336000
f1_keywords:
- vbaof11.chm336000
ms.prod: office
api_name:
- Office.PickerProperty
ms.assetid: fd3702fe-bf03-f22c-78c2-ac6c47a1d028
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerProperty object (Office)

Represents an object for passing a custom property.


## Example

The following code sets the **[PickerDialog](Office.PickerDialog.md)** properties and then displays the **PickerDialog**.


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

- [PickerProperty object members](overview/Library-Reference/pickerproperty-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]