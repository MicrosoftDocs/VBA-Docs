---
title: PickerResults.Item property (Office)
keywords: vbaof11.chm339001
f1_keywords:
- vbaof11.chm339001
ms.prod: office
api_name:
- Office.PickerResults.Item
ms.assetid: a871f16f-8278-87bd-3633-5e273bc8865f
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerResults.Item property (Office)

Retrieves a **[PickerResult](Office.PickerResult.md)** object at the specified index. Read-only.


## Syntax

_expression_.**Item**(_Index_)

_expression_ An expression that returns a **[PickerResults](Office.PickerResults.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Integer**|Specifies an integer representing the indexed location of the **PickerResult** object.|

## Example

The following code displays the **[PickerDialog](Office.PickerDialog.md)** user interface, gets results, and then enumerates those results.


```vb
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
Dim objPickerProperty As PickerProperty 
Dim objPickerExistingResults As PickerResults 
Dim objPickerExistingResults As PickerResult 
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
 
' Enumerate the results. 
For index = 1 To objPickerResults.Count-1 
 Debug.Print objPickerResults.Item(index).Id 
 Debug.Print objPickerResults.Item(index).DisplayName 
 Debug.Print objPickerResults.Item(index).Type 
 Debug.Print objPickerResults.Item(index).SIPId 
Next
```


## See also

- [PickerResults object members](overview/Library-Reference/pickerresults-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]