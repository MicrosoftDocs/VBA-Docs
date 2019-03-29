---
title: ListDataFormat object (Excel)
keywords: vbaxl10.chm757072
f1_keywords:
- vbaxl10.chm757072
ms.prod: excel
api_name:
- Excel.ListDataFormat
ms.assetid: d972f320-6865-a684-0f46-8c34b2eea482
ms.date: 03/30/2019
localization_priority: Normal
---


# ListDataFormat object (Excel)

The **ListDataFormat** object holds all the data type properties of the **[ListColumn](Excel.ListColumn.md)** object. These properties are read-only.


## Remarks

Use the **ListDataFormat** property of the **ListColumn** object to return a **ListDataFormat** object. The default property of the **ListDataFormat** object is the **[Type](Excel.ListDataFormat.Type.md)** property, which indicates the data type of the list column. This allows the user to write code without specifying the **Type** property.

<!--can't find a ListDataFormat property to link to-->

## Example

The following code example creates a linked list from a SharePoint list. It then checks to see if field 2 is required (field 1 is the ID field, which is read-only). If it's a required text field, the same data is written in all existing records.

> [!NOTE] 
> The following code example assumes that you will substitute a valid server name and the list GUID in the variables _strServerName_ and _strListGUID_. Additionally, the server name must be followed by `"/_vti_bin"` or the sample will not work.


```vb
Dim objListObject As ListObject 
Dim objDataRange As Range 
Dim strListGUID as String 
Dim strServerName as String 
 
strServerName = "https://<servername>/_vti_bin" 
strListGUID = "{<listguid>}" 
 
Set objListObject = Sheet1.ListObjects.Add(xlSrcExternal, _ 
 Array(strServerName, strListGUID), True, xlYes, Range("A1")) 
 
With objListObject.ListColumns(2) 
 Set objDataRange = .Range.Offset(1, 0).Resize(.Range.Rows.Count - 2, 1) 
 If .ListDataFormat.Type = xlListDataTypeText And .ListDataFormat.Required Then 
 objDataRange.Value = "Hello World" 
 End If 
End With 

```

## Properties

- [AllowFillIn](Excel.ListDataFormat.AllowFillIn.md)
- [Application](Excel.ListDataFormat.Application.md)
- [Choices](Excel.ListDataFormat.Choices.md)
- [Creator](Excel.ListDataFormat.Creator.md)
- [DecimalPlaces](Excel.ListDataFormat.DecimalPlaces.md)
- [DefaultValue](Excel.ListDataFormat.DefaultValue.md)
- [IsPercent](Excel.ListDataFormat.IsPercent.md)
- [lcid](Excel.ListDataFormat.lcid.md)
- [MaxCharacters](Excel.ListDataFormat.MaxCharacters.md)
- [MaxNumber](Excel.ListDataFormat.MaxNumber.md)
- [MinNumber](Excel.ListDataFormat.MinNumber.md)
- [Parent](Excel.ListDataFormat.Parent.md)
- [ReadOnly](Excel.ListDataFormat.ReadOnly.md)
- [Required](Excel.ListDataFormat.Required.md)
- [Type](Excel.ListDataFormat.Type.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]