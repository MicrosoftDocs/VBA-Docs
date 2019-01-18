---
title: PickerResult object (Office)
keywords: vbaof11.chm338000
f1_keywords:
- vbaof11.chm338000
ms.prod: office
api_name:
- Office.PickerResult
ms.assetid: 5229d2ad-a32e-a864-9de4-dc651199ff58
ms.date: 06/08/2017
localization_priority: Normal
---


# PickerResult object (Office)

Represents a resolved or selected item of data.


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
|[Application](Office.PickerResult.Application.md)|
|[Creator](Office.PickerResult.Creator.md)|
|[DisplayName](Office.PickerResult.DisplayName.md)|
|[DuplicateResults](Office.PickerResult.DuplicateResults.md)|
|[Fields](Office.PickerResult.Fields.md)|
|[Id](Office.PickerResult.Id.md)|
|[ItemData](Office.PickerResult.ItemData.md)|
|[SIPId](Office.PickerResult.SIPId.md)|
|[SubItems](Office.PickerResult.SubItems.md)|
|[Type](Office.PickerResult.Type.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
