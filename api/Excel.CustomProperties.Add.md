---
title: CustomProperties.Add method (Excel)
keywords: vbaxl10.chm680073
f1_keywords:
- vbaxl10.chm680073
ms.prod: excel
api_name:
- Excel.CustomProperties.Add
ms.assetid: 11165b03-e459-51c4-505f-67260ab8aaf9
ms.date: 04/23/2019
localization_priority: Normal
---


# CustomProperties.Add method (Excel)

Adds custom property information.


## Syntax

_expression_.**Add** (_Name_, _Value_)

_expression_ A variable that represents a **[CustomProperties](Excel.CustomProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the custom property.|
| _Value_|Required| **Variant**|The value of the custom property.|

## Return value

A **[CustomProperty](Excel.CustomProperty.md)** object that represents the custom property information.


## Example

This example adds identifier information to the active worksheet and returns the name and value to the user.

```vb
Sub CheckCustomProperties() 
 
 Dim wksSheet1 As Worksheet 
 
 Set wksSheet1 = Application.ActiveSheet 
 
 ' Add metadata to worksheet. 
 wksSheet1.CustomProperties.Add _ 
 Name:="Market", Value:="Nasdaq" 
 
 ' Display metadata. 
 With wksSheet1.CustomProperties.Item(1) 
 MsgBox .Name & vbTab & .Value 
 End With 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]