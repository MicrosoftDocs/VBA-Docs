---
title: PickerDialog object (Office)
keywords: vbaof11.chm340000
f1_keywords:
- vbaof11.chm340000
ms.prod: office
api_name:
- Office.PickerDialog
ms.assetid: 279b1a6a-f09d-a0e7-89c9-aac6c581439f
ms.date: 11/12/2018
localization_priority: Normal
---


# PickerDialog object (Office)

Provides dialog user interface functionality for picking people or picking data.

## Remarks

Get the **PickerDialog** object through the **PickerDialog** property in **Application** object.


## Example

The following code sets the Picker Dialog properties and then displays the Picker Dialog.


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


## Methods

|Name|Description|
|:---|:----------|
|[CreatePickerResults](Office.PickerDialog.CreatePickerResults.md)|Creates an empty **PickerResults** object.|
|[Resolve](Office.PickerDialog.Resolve.md)|Resolves the token using the Picker Dialog and retrieves the results. |
|[Show](Office.PickerDialog.Show.md)|Displays the Picker Dialog with already specified data handler and given options. |

## Properties

|Name|Description|
|:---|:----------|
|[Application](Office.PickerDialog.Application.md)|Gets an **Application** object that represents the container application for the **PickerDialog** object. Read-only. |
|[Creator](Office.PickerDialog.Creator.md)|Gets a 32-bit integer that indicates the application in which the **PickerDialog** object was created. Read-only. |
|[DataHandlerId](Office.PickerDialog.DataHandlerId.md)|Sets or gets the GUID of the Picker Dialog data handler component. Read/write. |
|[Properties](Office.PickerDialog.Properties.md)|Returns the **PickerProperties** object to specify custom properties for data handler component. Read-only. |
|[Title](Office.PickerDialog.Title.md)|Sets or returns the title of a picker dialog displayed in the Picker Dialog. Read/write. |

## See also

- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)
- [PickerDialog interface](https://docs.microsoft.com/dotnet/api/microsoft.office.core.pickerdialog?view=office-pia)
- [Object Picker dialog box interfaces](https://docs.microsoft.com/windows/desktop/ad/object-picker-dialog-box-interfaces)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]