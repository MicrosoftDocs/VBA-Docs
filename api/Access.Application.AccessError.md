---
title: Application.AccessError method (Access)
keywords: vbaac10.chm12556
f1_keywords:
- vbaac10.chm12556
ms.prod: access
api_name:
- Access.Application.AccessError
ms.assetid: 811ef090-bdd4-5d1d-afc5-782470f57483
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.AccessError method (Access)

You can use the **AccessError** method to return the descriptive string associated with a Microsoft Access or Data Access Objects (DAO) error.


## Syntax

_expression_.**AccessError** (_ErrorNumber_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ErrorNumber_|Required|**Variant**|The number of the error for which you wish to return a descriptive string.|

## Return value

Variant


## Remarks

You can use the **AccessError** method to return the descriptive string associated with a Microsoft Access or DAO error when the error hasn't actually occurred, but you cannot use it for ADO errors.

You can use the Visual Basic **Raise** method to raise a Visual Basic error. After you've raised the error, you can determine its associated descriptive string by reading the **Description** property of the **Err** object.

You can't use the **Raise** method to raise a Microsoft Access or DAO error. However, you can use the **AccessError** method to return the descriptive string associated with these errors, without having to generate the error.

You can use the **AccessError** method to return a descriptive string from within a form's **Error** event.

If the Microsoft Access error has occurred, you can return the descriptive string by using either the **AccessError** method or the **Description** property of the Visual Basic **Err** object.


## Example

The following function returns an error string for any valid error number.

> [!NOTE] 
> You must have your error trapping options set to **Break on Unhandled Errors** for the code to run in the VBA IDE. You can set this option on the **General** tab of the **Options** dialog box found on the VBA Tools menu.


```vb
Function ErrorString(ByVal lngError As Long) As String 
 
 Const conAppError = "Application-defined or " & _ 
 "object-defined error" 
 
 On Error Resume Next 
 Err.Raise lngError 
 
 If Err.Description = conAppError Then 
 ErrorString = AccessError(lngError) 
 ElseIf Err.Description = vbNullString Then 
 MsgBox "No error string associated with this number." 
 Else 
 ErrorString = Err.Description 
 End If 
 
End Function
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]