---
title: Folder.GetTable method (Outlook)
keywords: vbaol11.chm2018
f1_keywords:
- vbaol11.chm2018
ms.prod: outlook
api_name:
- Outlook.Folder.GetTable
ms.assetid: 08d184cb-0c41-01b1-abc5-305476380f8b
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.GetTable method (Outlook)

Obtains a **[Table](Outlook.Table.md)** object that contains items filtered by _Filter_.


## Syntax

_expression_. `GetTable`( `_Filter_` , `_TableContents_` )

_expression_ A variable that represents a '[Folder](Outlook.Folder.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Filter_|Optional| **String**|A filter in Microsoft Jet or DAV Searching and Locating (DASL) syntax that specifies the criteria for items in the parent  **Folder**.|
| _TableContents_|Optional| **[OlTableContents](Outlook.OlTableContents.md)**|Specifies the type of items in the folder that  **GetTable** returns. The default is **olUserItems**.|

## Return value

A **Table** that contains items in the parent **[Folder](Outlook.Folder.md)** that meet the criteria in _Filter_. By default, _TableContents_ is **olUserItems** and the returned **Table** contains only the filtered items that are not hidden.


## Remarks

If  _Filter_ is a blank string or the _Filter_ parameter is omitted, **GetTable** returns a **Table** with rows representing all the items in the **Folder**. If _Filter_ is a blank string or the _Filter_ parameter is omitted and _TableContents_ is **olHiddenItems**, **GetTable** returns a **Table** with rows representing all the hidden items in the **Folder**.

For more information on filters, see [Filtering Items](../outlook/How-to/Search-and-Filter/filtering-items.md) and [Referencing Properties by Namespace](../outlook/How-to/Navigation/referencing-properties-by-namespace.md).

 **GetTable** returns a **Table** with the default column set for the folder type of the parent **Folder**. To modify the default column set, use the **[Add](Outlook.Columns.Add.md)**, **[Remove](Outlook.Columns.Remove.md)**, or **[RemoveAll](Outlook.Columns.RemoveAll.md)** methods of the **[Columns](Outlook.Columns.md)** collection object. When _TableContents_ is **olHiddenItems**, the default column set is always the default column set for a mail folder even though the parent **Folder** might be, for example, a Contacts folder. For more information on default column sets, see [Default Properties Displayed in a Table Object](../outlook/How-to/Search-and-Filter/default-properties-displayed-in-a-table-object.md).

You can use  **[Table.Restrict](Outlook.Table.Restrict.md)** to apply subsequent filters to a **Table** that is based on the **Folder** object.


## Example

The following code sample illustrates how to use  **Folder.GetTable** to obtain a **Table** object based on the **LastModificationTime** of items in the Inbox. It then enumerates and prints the values of a couple of default properties of these items.


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
  
    'Enumerate the table using test for EndOfTable  
    Do Until (oTable.EndOfTable)  
        Set oRow = oTable.GetNextRow()  
        Debug.Print (oRow("Subject"))  
        Debug.Print (oRow("LastModificationTime"))  
    Loop  
End Sub
```


## See also


[Folder Object](Outlook.Folder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]