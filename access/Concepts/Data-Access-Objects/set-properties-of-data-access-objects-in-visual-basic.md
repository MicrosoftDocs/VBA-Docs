---
title: Set properties of Data Access Objects in Visual Basic
keywords: vbaac10.chm5188066
f1_keywords:
- vbaac10.chm5188066
ms.prod: access
ms.assetid: 8942307f-950d-f39d-cab2-ba4fa387b438
ms.date: 09/21/2018
localization_priority: Normal
---


# Set properties of Data Access Objects in Visual Basic

**Applies to:** Access 2013 | Access 2016

Data Access Objects (DAO) enable you to manipulate the structure of your database and the data it contains from Visual Basic. Many DAO objects correspond to objects that you see in your database—for example, a **TableDef** object corresponds to a Microsoft Access table. A **Field** object corresponds to a field in a table.

Most of the properties you can set for DAO objects are DAO properties. These properties are defined by the Microsoft Access database engine and are set the same way in any application that includes the Access database engine. Some properties that you can set for DAO objects are defined by Microsoft Access, and aren't automatically recognized by the Access database engine. How you set properties for DAO objects depends on whether a property is defined by the Access database engine or by Microsoft Access.

## Set DAO properties for DAO objects

To set a property that's defined by the Access database engine, refer to the object in the DAO hierarchy. The easiest and fastest way to do this is to create object variables that represent the different objects you need to work with, and refer to the object variables in subsequent steps in your code. For example, the following code creates a new **TableDef** object and sets its **Name** property:


```vb
Dim dbs As DAO.Database 
Dim tdf As DAO.TableDef 
Set dbs = CurrentDb 
Set tdf = dbs.CreateTableDef 
tdf.Name = "Contacts"
```


## Set Microsoft Access properties for DAO objects

When you set a property that's defined by Microsoft Access, but applies to a DAO object, the Access database engine doesn't automatically recognize the property as a valid property. The first time you set the property, you must create the property and append it to the **Properties** collection of the object to which it applies. Once the property is in the **Properties** collection, it can be set in the same manner as any DAO property.

If the property is set for the first time in the user interface, it's automatically added to the **Properties** collection, and you can set it normally.

When writing procedures to set properties defined by Microsoft Access, you should include error-handling code to verify that the property you are setting already exists in the **Properties** collection. See the Help topic about the **CreateProperty** method or the individual property topic for more information.

Keep in mind that when you create the property, you must correctly specify its **Type** property before you append it to the **Properties** collection. You can determine the **Type** property based on the information in the Settings section of the Help topic for the individual property. The following table provides some guidelines for determining the setting of the **Type** property.

|**If the property setting is**|**The Type property setting should be**|
|:-----|:-----|
|A string|**dbText**|
|**True** / **False**|**dbBoolean**|
|An integer|**dbInteger**|

<br/>

The following table lists some Microsoft Access-defined properties that apply to DAO objects.

|**DAO object**|**Microsoft Access-defined properties**|
|:-----|:-----|
|**Database**|**AppTitle**, **AppIcon**, **StartupShowDBWindow**, **StartupShowStatusBar**, **AllowShortcutMenus**, **AllowFullMenus**, **AllowBuiltInToolbars**, **AllowToolbarChanges**, **AllowBreakIntoCode**, **AllowSpecialKeys**, **Replicable**, **ReplicationConflictFunction**|
|SummaryInfo **Container**|**Title**, **Subject**, **Author**, **Manager**, **Company**, **Category**, **Keywords**, **Comments**, **Hyperlink Base** (See the **Summary** tab of the **_DatabaseName_ Properties** dialog box, available by selecting **Database Properties** on the **File** menu.)|
|UserDefined **Container**|(See the **Summary** tab of the **_DatabaseName_ Properties** dialog box, available by selecting **Database Properties** on the **File** menu.)|
|**TableDef**|**DatasheetBackColor**, **DatasheetCellsEffect**, **DatasheetFontHeight**, **DatasheetFontItalic**, **DatasheetFontName**, **DatasheetFontUnderline**, **DatasheetFontWeight**, **DatasheetForeColor**, **DatasheetGridlinesBehavior**, **DatasheetGridlinesColor**, **Description**, **FrozenColumns**, **RowHeight**, **ShowGrid**|
|**QueryDef**|**DatasheetBackColor**, **DatasheetCellsEffect**, **DatasheetFontHeight**, **DatasheetFontItalic**, **DatasheetFontName**, **DatasheetFontUnderline**, **DatasheetFontWeight**, **DatasheetForeColor**, **DatasheetGridlinesBehavior**, **DatasheetGridlinesColor**, **Description**, **FailOnError**, **FrozenColumns**, **LogMessages**, **MaxRecords**, **RecordLocks**, **RowHeight**, **ShowGrid, UseTransaction**|
|**Field**|**Caption**, **ColumnHidden**, **ColumnOrder**, **ColumnWidth**, **DecimalPlaces**, **Description**, **Format**, **InputMask**|

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
