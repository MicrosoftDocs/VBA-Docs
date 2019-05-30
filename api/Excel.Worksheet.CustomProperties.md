---
title: Worksheet.CustomProperties property (Excel)
keywords: vbaxl10.chm175152
f1_keywords:
- vbaxl10.chm175152
ms.prod: excel
api_name:
- Excel.Worksheet.CustomProperties
ms.assetid: 49862772-caff-90a1-3266-c8b158003aff
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.CustomProperties property (Excel)

Returns a **[CustomProperties](Excel.CustomProperties.md)** object representing the identifier information associated with a worksheet.


## Syntax

_expression_.**CustomProperties**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

For the **CustomProperties** property, identifier information for a worksheet can represent metadata for XML.


## Example

In this example, Microsoft Excel adds identifier information to the active worksheet and returns the name and value to the user.

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