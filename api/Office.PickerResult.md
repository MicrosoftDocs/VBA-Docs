---
title: PickerResult object (Office)
keywords: vbaof11.chm338000
f1_keywords:
- vbaof11.chm338000
ms.prod: office
api_name:
- Office.PickerResult
ms.assetid: 5229d2ad-a32e-a864-9de4-dc651199ff58
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerResult object (Office)

Represents a resolved or selected item of data.


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

- [PickerResult object members](overview/Library-Reference/pickerresult-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]