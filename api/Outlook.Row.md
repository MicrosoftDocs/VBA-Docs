---
title: Row object (Outlook)
keywords: vbaol11.chm3167
f1_keywords:
- vbaol11.chm3167
ms.prod: outlook
api_name:
- Outlook.Row
ms.assetid: 06db3fa4-1649-48bf-3b86-ffdf99a47305
ms.date: 06/08/2017
localization_priority: Normal
---


# Row object (Outlook)

Represents a row of data in the  **[Table](Outlook.Table.md)** object.


## Remarks

A **Table** is composed of rows and columns. It represents a read-only dynamic rowset of data in a **[Folder](Outlook.Folder.md)** or **[Search](Outlook.Search.md)** object. You can regard each row of a **Table** as an item in the folder, and each column as a property of the item. By default, the **Table** contains only a subset of properties for items in the folder. This makes the **Table** an in-memory lightweight rowset that supports fast enumeration and filtering of items in the folder.

 If the **Table** object is obtained from **[Folder.GetTable](Outlook.Folder.GetTable.md)**, you can further specify a filter (in **[Table.Restrict](Outlook.Table.Restrict.md)**) to obtain a more restricted set of rows in the **Table**.

 You can use the Table methods: **[FindRow](Outlook.Table.FindRow.md)**, **[FindNextRow](Outlook.Table.FindNextRow.md)**, **[GetNextRow](Outlook.Table.GetNextRow.md)**, and **[MoveToStart](Outlook.Table.MoveToStart.md)** to obtain a specific row in a **Table**.

 Use **[Row.GetValues](Outlook.Row.GetValues.md)** to obtain an array of values that correspond to column values at that row in the **Table**.

 Use the helper functions **[Row.BinaryToString](Outlook.Row.BinaryToString.md)**, **[Row.LocalTimeToUTC](Outlook.Row.LocalTimeToUTC.md)**, and **[Row.UTCToLocalTime](Outlook.Row.UTCToLocalTime.md)** to facilitate type conversion of column values at a specific row. For more information on property value representation in a **Table**, see [Factors Affecting Property Value Representation in the Table and View Classes](../outlook/How-to/Search-and-Filter/factors-affecting-property-value-representation-in-the-table-and-view-classes.md).

 Although additions and deletions of the underlying folder are reflected by the rows in the **Table**, the **Table** does not support any events for adding, changing, and removing of rows. If you require a writeable object from the **Table** row, obtain the Entry ID for that row from the default EntryID column in the **Table** and then use the **[GetItemFromID](Outlook.NameSpace.GetItemFromID.md)** method of the **[NameSpace](Outlook.NameSpace.md)** object to obtain a full item, such as a **[MailItem](Outlook.MailItem.md)** or **[ContactItem](Outlook.ContactItem.md)**, that supports read-write operations. For more information on default columns in a **Table**, see [Default Properties Displayed in a Table Object](../outlook/How-to/Search-and-Filter/default-properties-displayed-in-a-table-object.md).


## Example

The following code sample illustrates how to obtain a **Table** object based on the **LastModificationTime** of items in the Inbox. It also shows how to customize columns in the **Table**, and how to enumerate and print the values of the corresponding properties of these items.


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
|[BinaryToString](Outlook.Row.BinaryToString.md)|
|[GetValues](Outlook.Row.GetValues.md)|
|[Item](Outlook.Row.Item.md)|
|[LocalTimeToUTC](Outlook.Row.LocalTimeToUTC.md)|
|[UTCToLocalTime](Outlook.Row.UTCToLocalTime.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Row.Application.md)|
|[Class](Outlook.Row.Class.md)|
|[Parent](Outlook.Row.Parent.md)|
|[Session](Outlook.Row.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]