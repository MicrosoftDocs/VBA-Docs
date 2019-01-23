---
title: PickerProperties.Add method (Office)
keywords: vbaof11.chm337003
f1_keywords:
- vbaof11.chm337003
ms.prod: office
api_name:
- Office.PickerProperties.Add
ms.assetid: a52c9607-1b0a-c37e-a3af-dc0550c64deb
ms.date: 01/22/2019
localization_priority: Normal
---


# PickerProperties.Add method (Office)

Adds a **[PickerProperty](Office.PickerProperty.md)** object to the collection.


## Syntax

_expression_.**Add** (_Id_, _Value_, _Type_)

_expression_ An expression that returns a **[PickerProperties](Office.PickerProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**String**|Key name of the property.|
| _Value_|Required|**String**|Value of the property.|
| _Type_|Required|**[MsoPickerField](office.msopickerfield.md)**|Type of the property.|

## Return value

PickerProperty


## Example

The following code sets various properties of the **[PickerDialog](office.pickerdialog.md)** object.


```vb
Dim objPickerDialog As PickerDialog 
Dim objPickerProperties As PickerProperties 
 
' Configure Picker Dialog properties. 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
Set objPickerProperties = objPickerDialog.Properties 
Set objPickerProperty = objPickerProperties.Add("SiteUrl", "https://my", msoPickerFieldtypeText) 

```


## See also

- [PickerProperties object members](overview/Library-Reference/pickerproperties-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]