---
title: Row.Item method (Outlook)
keywords: vbaol11.chm2245
f1_keywords:
- vbaol11.chm2245
ms.prod: outlook
api_name:
- Outlook.Row.Item
ms.assetid: fa9a6b26-ddfe-f306-5f45-140756f398c9
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Item method (Outlook)

Obtains an  **Object** that represents the value for the **[Row](Outlook.Row.md)** object at the column specified by _Index_.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [Row](Outlook.Row.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a  **Long** representing the column index for the **[Columns](Outlook.Columns.md)** collection or a **String** representing the **[Name](Outlook.Column.Name.md)** of the **[Column](Outlook.Column.md)**.|

## Return value

A  **Variant** that represents the value of a property (as specified by _Index_ ) of an item (as specified by the parent **Row**).


## Remarks

The  **Item** method is the default method of the **Row** object, meaning that the method can be used implicitly. The following two lines of code both access the value of the **Subject** property at the specified **Row** in a **[Table](Outlook.Table.md)** :

 `Row.Item("Subject")`

 `Row("Subject")`

If a  **Column** has been added to a **Table** using a property name referencing a namespace, you must reference the **Column** in the **Row.Item** method by the same namespace reference. If you use an explicit built-in name reference in **Row.Item**, you will get an error.


## Example

The following code sample illustrates how to obtain a  **Table** object based on the **LastModificationTime** of items in the Inbox. It then enumerates and prints the values of a couple of default properties of these items. Since the **Item** method is the default method of the **Row** object, it uses the **Item** method in an implicit way.


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


[Row Object](Outlook.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]