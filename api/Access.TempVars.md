---
title: TempVars Object (Access)
keywords: vbaac10.chm14073
f1_keywords:
- vbaac10.chm14073
ms.prod: access
api_name:
- Access.TempVars
ms.assetid: aa81b18b-5e9f-ae44-cbcf-55cf6e37b7f6
ms.date: 06/08/2017
---


# TempVars Object (Access)

Represents the collection of  **[TempVar](Access.TempVar.md)** objects.


## Remarks

Use the  **[Add](./Access.TempVars.Add.md)** method or the[SetTempVar](http://msdn.microsoft.com/library/9c3b7bee-02c5-efbf-1276-4c4a1f7802d9%28Office.15%29.aspx) macro action to create a **TempVar** object.

Use the  **[Remove](./Access.TempVars.Remove.md)** method or the[RemoveTempVar](http://msdn.microsoft.com/library/409fd836-4a53-cefd-4264-8cee0fa8ac52%28Office.15%29.aspx) macro action to delete a **TempVar** object from the **TempVars** collection.

Use the  **[RemoveAll](./Access.TempVars.RemoveAll.md)** method or[RemoveAllTempVars](http://msdn.microsoft.com/library/409fd836-4a53-cefd-4264-8cee0fa8ac52%28Office.15%29.aspx) macro action to delete all **TempVar** objects from the **TempVars** collection.

The  **TempVars** collection can store up to 255 **TempVar** objects. If you do not remove a **TempVar** object, it will remain in memory until you close the database. It is a good practice to remove **TempVar** object variables when you are finished using them.

To refer to a  **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:


-  **TempVar** ![name]
    

## Methods



|**Name**|
|:-----|
|[Add](./Access.TempVars.Add.md)|
|[Remove](./Access.TempVars.Remove.md)|
|[RemoveAll](./Access.TempVars.RemoveAll.md)|

## Properties



|**Name**|
|:-----|
|[Application](./Access.TempVars.Application.md)|
|[Count](./Access.TempVars.Count.md)|
|[Item](./Access.TempVars.Item.md)|
|[Parent](./Access.TempVars.Parent.md)|

## See also


[Access Object Model Reference](./overview/object-model-access-vba-reference.md)
[TempVars Object Members](http://msdn.microsoft.com/library/5c83c870-c66c-8fd9-0ac6-06766b14a6fc%28Office.15%29.aspx)
