---
title: CustomProperties object (Excel)
keywords: vbaxl10.chm679072
f1_keywords:
- vbaxl10.chm679072
ms.prod: excel
api_name:
- Excel.CustomProperties
ms.assetid: f0f38570-e3bf-58ad-ab8a-e412ad869907
ms.date: 03/29/2019
localization_priority: Normal
---


# CustomProperties object (Excel)

A collection of **[CustomProperty](excel.customproperty.md)** objects that represents additional information. The information can be used as metadata for XML.


## Remarks

Use the **[CustomProperties](Excel.Worksheet.CustomProperties.md)** property of the **Worksheet** object to return a **CustomProperties** collection.

After a **CustomProperties** collection is returned, you can add metadata to worksheets and perform additional actions depending on which you choose to work with.

To add metadata to a worksheet, use the **CustomProperties** property with the **Add** method.


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


## Methods

- [Add](Excel.CustomProperties.Add.md)

## Properties

- [Application](Excel.CustomProperties.Application.md)
- [Count](Excel.CustomProperties.Count.md)
- [Creator](Excel.CustomProperties.Creator.md)
- [Item](Excel.CustomProperties.Item.md)
- [Parent](Excel.CustomProperties.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]