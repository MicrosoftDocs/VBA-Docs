---
title: SetObjectOwner method (ADOX)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 22c5d2d9-c7b2-3c3a-0b1f-a2e5bc46395c
ms.date: 06/08/2017
localization_priority: Normal
---


# SetObjectOwner method (ADOX)

**Applies to:** Access 2013 | Access 2016

Specifies the owner of an object in a **Catalog**.

## Parameters

-  _ObjectName_
    
    - A **String** value that specifies the name of the object for which to specify the owner.  
    
-  _ObjectType_
    
    - A **Long** value that can be one of the **ObjectTypeEnum** constants that specifies the owner type.
    
-  _OwnerName_
    
    - A **String** value that specifies the **Name** of the **User** or **Group** to own the object.
    
-  _ObjectTypeId_
    
    - Optional. A **Variant** value that specifies the GUID for a provider object type not defined by the OLE DB specification. This parameter is required if _ObjectType_ is set to **adPermObjProviderSpecific**; otherwise, it is not used.
    

## Remarks

An error will occur if the provider does not support specifying object owners.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]