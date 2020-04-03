---
title: Table.Sort method (Outlook)
keywords: vbaol11.chm2235
f1_keywords:
- vbaol11.chm2235
ms.prod: outlook
api_name:
- Outlook.Table.Sort
ms.assetid: 4e4867c2-27b8-f920-59ce-b60116d22054
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Sort method (Outlook)

Sorts the rows of the  **[Table](Outlook.Table.md)** by the property specified in _SortProperty_ and resets the current row to just before the first row in the **Table**.


## Syntax

_expression_.**Sort** (_SortProperty_, _Descending_)

_expression_ A variable that represents a [Table](Outlook.Table.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SortProperty_|Required| **String**|Specifies the property to use to sort the rows of the  **Table**.|
| _Descending_|Optional| **Boolean**|Whether to sort the  **Table** in descending order.|

## Remarks

 _SortProperty_ can be any explicit built-in property or custom property, with the exception of binary and multi-valued properties. The property must be referenced by its explicit string name; it cannot be referenced by namespace. For further information on specifying sort properties, see [Sorting Items in a Folder](../outlook/How-to/Items-Folders-and-Stores/sorting-items-in-a-folder.md).

Sorting the table is equivalent to calling a **[MoveToStart](Outlook.Table.MoveToStart.md)** method. The cursor will be positioned to the start of the Table.

If  **Table.Sort** and then **[Table.Restrict](Outlook.Table.Restrict.md)** are called, the filtered items in the new **Table** will be sorted by the same _SortProperty_ and _SortOrder_.

 **Table.Sort** only supports sorting on a single column.


## Example

The following code sample shows how to sort the rows in a **Table** based on the **ReceivedTime** property, and prints the value of the MAPI normalized subject property for each row in the sorted table.


```vb
Sub SortTableByReceivedTime() 
 
 Dim oT As Outlook.Table 
 
 Dim oRow As Outlook.Row 
 
 Set oT = Session.GetDefaultFolder(olFolderInbox).GetTable 
 
 'Add normalized subject (subject without RE:, FW: and other prefixes)to the column set 
 
 oT.Columns.Add ("http://schemas.microsoft.com/mapi/proptag/0x0E1D001E") 
 
 
 
 'Sort by ReceivedTime in descending order 
 
 oT.Sort "[ReceivedTime]", True 
 
 
 
 Do Until oT.EndOfTable 
 
 Set oRow = oT.GetNextRow 
 
 'Print the normalized subject of each row 
 
 Debug.Print oRow("http://schemas.microsoft.com/mapi/proptag/0x0E1D001E") 
 
 Loop 
 
End Sub
```


## See also


[Table Object](Outlook.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]