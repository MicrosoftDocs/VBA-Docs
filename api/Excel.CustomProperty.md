---
title: CustomProperty object (Excel)
keywords: vbaxl10.chm681072
f1_keywords:
- vbaxl10.chm681072
ms.prod: excel
api_name:
- Excel.CustomProperty
ms.assetid: df8b58d8-ccfd-00bb-723a-a9c328f0b38b
ms.date: 03/29/2019
localization_priority: Normal
---


# CustomProperty object (Excel)

Represents identifier information, which can be used as metadata for XML.


## Remarks

Use the **Add** method or the **Item** property of the **[CustomProperties](Excel.CustomProperties.md)** collection to return a **CustomProperty** object.

After a **CustomProperty** object is returned, you can add metadata to worksheets by using the **[CustomProperties](Excel.Worksheet.CustomProperties.md)** property of the **Worksheet** object with the **Add** method.


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


## Methods

- [Delete](Excel.CustomProperty.Delete.md)

## Properties

- [Application](Excel.CustomProperty.Application.md)
- [Creator](Excel.CustomProperty.Creator.md)
- [Name](Excel.CustomProperty.Name.md)
- [Parent](Excel.CustomProperty.Parent.md)
- [Value](Excel.CustomProperty.Value.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]