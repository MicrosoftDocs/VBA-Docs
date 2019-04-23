---
title: PickerResults.Add method (Office)
keywords: vbaof11.chm339003
f1_keywords:
- vbaof11.chm339003
ms.prod: office
api_name:
- Office.PickerResults.Add
ms.assetid: cf6e4f0f-4373-3caa-ddb3-512ca5c4675f
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerResults.Add method (Office)

Adds a **[PickerResult](Office.PickerResult.md)** object to the **PickerResults** collection.


## Syntax

_expression_.**Add** (_Id_, _DisplayName_, _Type_, _SIPId_, _ItemData_, _SubItems_)

_expression_ An expression that returns a **[PickerResults](Office.PickerResults.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**String**|Represents an identifier of the **PickerResult**.|
| _DisplayName_|Required|**String**|Represents a display name of the **PickerResult**. |
| _Type_|Required|**String**|Represents a type of the **PickerResult**.|
| _SIPId_|Optional|**String**|Currently not supported. The **SIPId** is the identifier for Office Communication Server. It is used only for the people picking scenario.|
| _ItemData_|Optional|**Variant**|Non-displaying item binding data.|
| _SubItems_|Optional|**Variant**|Displays the purpose or non-display purpose field data of the **PickerResult**. It is used for passing column values in the **PickerDialog**.|

## Return value

PickerResult


## Example

The following code sets the **[PickerDialog](Office.PickerDialog.md)** properties and then displays the **PickerDialog**.


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
```


## See also

- [PickerResults object members](overview/Library-Reference/pickerresults-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]