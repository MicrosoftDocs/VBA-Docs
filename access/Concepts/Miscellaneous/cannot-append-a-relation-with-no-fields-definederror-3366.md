---
title: Cannot append a relation with no fields defined. (Error 3366)
keywords: jeterr40.chm5003366
f1_keywords:
- jeterr40.chm5003366
ms.prod: access
ms.assetid: cac57d13-5705-c67a-2621-8076346a70a3
ms.date: 06/08/2017
localization_priority: Normal
---


# Cannot append a relation with no fields defined. (Error 3366)

  

**Applies to:** Access 2013 | Access 2016

You are trying to append a **Relation** object to a **Relations** collection, but the **Relation** object has no fields.

 To correctly append a Relation


1. Use the **CreateRelation** method to create the **Relation** object. Set the **Table**, **ForeignTable**, and **Attributes** properties of the **Relation** object, if you did not specify them as arguments to the **CreateRelation** method. Use the **CreateField** method to create a new **Field** object for each field in the primary and foreign keys of the relationship.
    
2. Set the **Name** (if you did not specify it as an argument to the **CreateField** method) and **ForeignName** properties of the **Field** object or objects to the corresponding **Name** property settings of the primary key and the foreign key **Field** objects of each field in the relationship.
    
3. Use the **Append** method to save the **Field** object or objects in the **Fields** collection of the **Relation** object.
    
4. Use the **Append** method to save the **Relation** object in the **Relations** collection of the database.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]