---
title: ListDataFormat.Choices property (Excel)
keywords: vbaxl10.chm758074
f1_keywords:
- vbaxl10.chm758074
ms.prod: excel
api_name:
- Excel.ListDataFormat.Choices
ms.assetid: c4a809e6-7977-28a1-1070-286e7df99409
ms.date: 04/30/2019
localization_priority: Normal
---


# ListDataFormat.Choices property (Excel)

Returns an **Array** of **String** values that contains the choices offered to the user by the **ListLookUp**, **ChoiceMulti**, and **Choice** data types of the **[DefaultValue](Excel.ListDataFormat.DefaultValue.md)** property. Read-only **Variant**.


## Syntax

_expression_.**Choices**

_expression_ A variable that represents a **[ListDataFormat](Excel.ListDataFormat.md)** object.


## Remarks

In Microsoft Excel, you cannot set any of the properties associated with the **ListDataFormat** object. However, you can set these properties by modifying the list on the server that is running Microsoft SharePoint Foundation.


## Example

The following example displays the setting of the **Choices** property for the third column in a list that is linked to a SharePoint list. In this example, it is assumed that the **DefaultValue** property has been set to the **Choice**, **ChoiceMulti**, or **ListLookup** data type.

```vb
Sub PrintChoices() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.Choices 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]