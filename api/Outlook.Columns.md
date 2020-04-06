---
title: Columns object (Outlook)
keywords: vbaol11.chm3190
f1_keywords:
- vbaol11.chm3190
ms.prod: outlook
api_name:
- Outlook.Columns
ms.assetid: 628bf0cf-4ee8-5e5c-09d7-89d7adf256ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns object (Outlook)

Represents the collection of  **[Column](Outlook.Column.md)** objects in a **[Table](Outlook.Table.md)** object.


## Remarks

The  **Columns** object supports enumerating **Column** objects in the **[Columns](Outlook.Columns.md)** collection object. It supports the COM interface, **IEnumerable**.


## Example

The following code sample illustrates how to obtain a  **Table** object based on the **LastModificationTime** of items in the Inbox. It also shows how to remove the default columns of the **Table**, add specific columns, and print the values of the corresponding properties of these items.


```vb
Sub RemoveAllAndAddColumns() 
 
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
|[Add](Outlook.Columns.Add.md)|
|[Item](Outlook.Columns.Item.md)|
|[Remove](Outlook.Columns.Remove.md)|
|[RemoveAll](Outlook.Columns.RemoveAll.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Columns.Application.md)|
|[Class](Outlook.Columns.Class.md)|
|[Count](Outlook.Columns.Count.md)|
|[Parent](Outlook.Columns.Parent.md)|
|[Session](Outlook.Columns.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]