---
title: Invalid operation. (Error 3219)
keywords: jeterr40.chm5003219
f1_keywords:
- jeterr40.chm5003219
ms.prod: access
ms.assetid: ab31a5dd-0979-2a03-3816-ef62ac370cae
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid operation. (Error 3219)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- You tried to write to a read-only property. See the Help topic for the property to determine whether it is read/write.
    
- You tried to use a method or property on a type of  **Recordset** object that the method or property does not apply to. See the Recordset object summary topic to determine which methods and properties apply to a given type of **Recordset** object.
    
- You tried to append a property to a  **Properties** collection of an object that does not support user-defined properties.
    
- You tried to use the  **Update** method on a read-only **Recordset** object.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]