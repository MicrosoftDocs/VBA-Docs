---
title: CustomProperty object (Excel)
keywords: vbaxl10.chm681072
f1_keywords:
- vbaxl10.chm681072
ms.prod: excel
api_name:
- Excel.CustomProperty
ms.assetid: df8b58d8-ccfd-00bb-723a-a9c328f0b38b
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomProperty object (Excel)

Represents identifier information. Identifier information can be used as metadata for XML.


## Remarks

Use the  **[Add](Excel.CustomProperties.Add.md)** method or the **[Item](Excel.CustomProperties.Item.md)** property of the **[CustomProperties](Excel.CustomProperties.md)** collection to return a **CustomProperty** object.

Once a  **CustomProperty** object is returned, you can add metadata to worksheets using the **[CustomProperties](Excel.Worksheet.CustomProperties.md)** property with the **Add** method.


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


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]