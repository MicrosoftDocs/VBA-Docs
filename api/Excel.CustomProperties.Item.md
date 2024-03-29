---
title: CustomProperties.Item property (Excel)
keywords: vbaxl10.chm680076
f1_keywords:
- vbaxl10.chm680076
api_name:
- Excel.CustomProperties.Item
ms.assetid: f2b9890b-2a25-e192-323b-dca72b461229
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# CustomProperties.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[CustomProperties](Excel.CustomProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**|The index number of the object.|

## Example

The following example demonstrates this feature. In this example, Microsoft Excel adds identifier information to the active worksheet and returns the name and value to the user.

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