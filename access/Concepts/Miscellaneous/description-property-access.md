---
title: Description property (Access)
ROBOTS: INDEX
keywords: vbaac10.chm3487
f1_keywords:
- vbaac10.chm3487
ms.prod: access
api_name:
- Access.Description
ms.assetid: b2933bc9-5e8b-9bee-d07b-2b015c530ebe
ms.date: 06/08/2017
localization_priority: Normal
---


# Description property (Access)

**Applies to:** Access 2013 | Access 2016

You can use the **Description** property to provide information about objects contained in the Database window as well as about individual table or query fields.


## Setting

For a database object, click **Properties** on the **View** menu and enter the description text in the **Description** box. For tables or queries, you can also enter the description in the table's or query's property sheet. An object's description appears next to the object's name in the Database window when you click **Details** on the **View** menu.

For individual table or query fields, enter the field description in the upper portion of table Design view or in the Field Properties property sheet in the Query window. The maximum length is 255 characters.

In Visual Basic , to set this property for the first time in a Microsoft Access project (.adp), you must create an application-defined property by using the **Add** method. In a Microsoft Access database (.mdb), you must use the DAO **CreateProperty** method.


## Remarks

An object's description is displayed in the Description column in the Details view of the Database window.

If you create controls by dragging a field from the field list, Microsoft Access copies the field's **Description** property to the control's **StatusBarText** property.

> [!NOTE] 
> For a linked table, Microsoft Access displays the connection information in the **Description** property.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
