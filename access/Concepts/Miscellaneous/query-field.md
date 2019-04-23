---
title: Query Field
ms.prod: access
ms.assetid: b39deb06-3814-ed9e-a3d5-558e3f3170e7
ms.date: 06/08/2017
localization_priority: Normal
---


# Query Field

  

**Applies to:** Access 2013 | Access 2016

A query field represents data from a table linked to the query. By default, a query field inherits all the properties that it has in the underlying table or query. For example, if a table design specifies the display format of the Order Date field as Medium Date in the field's  **Format** property, the Order Date field is formatted in the query recordset as Medium Date. Because the underlying field properties are the defaults, they aren't displayed on the property sheet.

If a field property is changed in the table design, the query field automatically inherits the change. If, however, you change a property within the query, the property setting within the table design is overridden. If the property is later changed in the table design, the change isn't reflected in the query.
You can set the properties for query fields within the Field Properties window of the query Design view.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]