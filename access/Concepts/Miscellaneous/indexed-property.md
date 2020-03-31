---
title: Indexed property
ROBOTS: INDEX
keywords: vbaac10.chm5187337
f1_keywords:
- vbaac10.chm5187337
ms.prod: access
api_name:
- Access.Indexed
ms.assetid: d410da7f-eb9d-5e03-06fa-b5756da357f4
ms.date: 06/08/2019
localization_priority: Normal
---


# Indexed property

**Applies to:** Access 2013 | Access 2016

You can use the **Indexed** property to set a single-field index. An index speeds up queries on the indexed fields as well as sorting and grouping operations. For example, if you search for specific employee names in a LastName field, you can create an index for this field to speed up the search for a specific name.

## Setting

The **Indexed** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|No|(Default) No index.|
|Yes (Duplicates OK)|The index allows duplicates.|
|Yes (No Duplicates)|The index doesn't allow duplicates.|

You can set this property only in the Field Properties section in table Design view. You can set a single-field index by setting the **Indexed** property in the Field Properties section in table Design view. You can set multiple-field indexes in the Indexes window. To open the **Indexes** window, on the **Design** tab, in the **Show/Hide** group, click **Indexes**.

If you add a single-field index in the Indexes window, Microsoft Access will set the **Indexed** property for the field to Yes.

In Visual Basic , use the ADO **Append** method of the **Indexes** collection to create an index for a field.


## Remarks

Use the **Indexed** property to find and sort records by using a single field in a table. The field can hold either unique or non-unique values. For example, you can create an index on an EmployeeID field in an Employees table in which each employee ID is unique or you can create an index on a Name field in which some names may be duplicates.

> [!NOTE] 
> You can't index Memo, Hyperlink, or OLE Object data type fields.

You can create as many indexes as you need. The indexes are created when you save the table and are automatically updated when you change or add records. You can add or delete indexes at any time in table Design view.

> [!TIP] 
> You can specify text that is commonly used at the beginning or the end of a field name (such as "ID", "code", or "num") for the **AutoIndex On Import/Create** option on the **Tables/Queries** tab, available by clicking **Options** on the **Tools** menu. When you import data files that contain this text in their field names, Microsoft Access creates an index for these fields.

If the primary key for a table is a single field, Microsoft Access will automatically set the **Indexed** property for that field to Yes (No Duplicates).

If you want to create multiple-field indexes, use the Indexes window.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
