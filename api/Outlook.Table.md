---
title: Table object (Outlook)
keywords: vbaol11.chm3166
f1_keywords:
- vbaol11.chm3166
ms.prod: outlook
api_name:
- Outlook.Table
ms.assetid: 0affaafd-93fe-227a-acee-e09a86cadc20
ms.date: 06/08/2017
localization_priority: Normal
---


# Table object (Outlook)

Represents a set of item data from a  **[Folder](Outlook.Folder.md)** or **[Search](Outlook.Search.md)** object, with items as rows of the table and properties as columns of the table.


## Remarks

The  **Table** represents a read-only dynamic rowset of data in a **Folder** or **Search** object. You can use **[Folder.GetTable](Outlook.Folder.GetTable.md)** or **[Search.GetTable](Outlook.Search.GetTable.md)** to obtain a **Table** object that represents a set of items in a folder or search folder. If the **Table** object is obtained from **Folder.GetTable**, you can further specify a filter (in **[Table.Restrict](Outlook.Table.Restrict.md)**) to obtain a subset of the items in the folder. If you do not specify any filter, you will obtain all the items in the folder.

By default, each item in the returned  **Table** contains only a default subset of its properties. You can regard each row of a **Table** as an item in the folder, each column as a property of the item, and the **Table** as an in-memory lightweight rowset that allows fast enumeration and filtering of items in the folder. Although additions and deletions of the underlying folder are reflected by the rows in the **Table**, the **Table** does not support any events for adding, changing, and removing of rows. If you require a writeable object from the **Table** row, obtain the Entry ID for that row from the default EntryID column in the **Table** and then use the **[GetItemFromID](Outlook.NameSpace.GetItemFromID.md)** method of the **[NameSpace](Outlook.NameSpace.md)** object to obtain a full item, such as a **[MailItem](Outlook.MailItem.md)** or **[ContactItem](Outlook.ContactItem.md)**, that supports read-write operations. For more information on default columns in a **Table**, see [Default Properties Displayed in a Table Object](../outlook/How-to/Search-and-Filter/default-properties-displayed-in-a-table-object.md).

 For more information on the **Table** object, see [Enumerating, Searching, and Filtering Items in a Folder](../outlook/How-to/Search-and-Filter/enumerating-searching-and-filtering-items-in-a-folder.md).


## Example

The following code sample illustrates how the  **Table** object can return a filtered set of items based on their **LastModificationTime** property. It also shows how to list the default properties as well as specific properties of the items.


```vb
Sub DemoTable() 
 
 'Declarations 
 
 Dim Filter As String 
 
 Dim oRow As Outlook.Row 
 
 Dim oTable As Outlook.Table 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 'Get a Folder object for the Inbox 
 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 
 'Restrict with Filter 
 
 Set oTable = oFolder.GetTable(Filter) 
 
 
 
 'Remove all columns in the default column set 
 
 oTable.Columns.RemoveAll 
 
 'Specify desired properties 
 
 With oTable.Columns 
 
 .Add ("Subject") 
 
 .Add ("LastModificationTime") 
 
 'PR_ATTR_HIDDEN referenced by the MAPI proptag namespace 
 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 
 End With 
 
 
 
 'Enumerate the table using test for EndOfTable 
 
 Do Until (oTable.EndOfTable) 
 
 Set oRow = oTable.GetNextRow() 
 
 Debug.Print (oRow("Subject")) 
 
 Debug.Print (oRow("LastModificationTime")) 
 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 
 Loop 
 
End Sub
```


## Methods



|Name|
|:-----|
|[FindNextRow](Outlook.Table.FindNextRow.md)|
|[FindRow](Outlook.Table.FindRow.md)|
|[GetArray](Outlook.Table.GetArray.md)|
|[GetNextRow](Outlook.Table.GetNextRow.md)|
|[GetRowCount](Outlook.Table.GetRowCount.md)|
|[MoveToStart](Outlook.Table.MoveToStart.md)|
|[Restrict](Outlook.Table.Restrict.md)|
|[Sort](Outlook.Table.Sort.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Table.Application.md)|
|[Class](Outlook.Table.Class.md)|
|[Columns](Outlook.Table.Columns.md)|
|[EndOfTable](Outlook.Table.EndOfTable.md)|
|[Parent](Outlook.Table.Parent.md)|
|[Session](Outlook.Table.Session.md)|

## See also


[Table Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]