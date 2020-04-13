---
title: Cannot perform this operation. (Error 3032)
keywords: jeterr40.chm5003032
f1_keywords:
- jeterr40.chm5003032
ms.prod: access
ms.assetid: 97a6b163-1ec8-176b-ee8d-d19610b29239
ms.date: 06/08/2019
localization_priority: Normal
---


# Cannot perform this operation. (Error 3032)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- You tried to delete the only user account in the group Admins. The group Admins must have at least one user account. If you want to delete this account, create a new account and add it to the group Admins, or add an existing account to the group Admins, and then delete the account.
    
- You tried to put a user in a group to which the user already belongs, by appending either a **Group** object to a **User** object's **Groups** collection that already has a **Group** object of the same name or a **User** object to a **Group** object's **Users** collection that already has a **User** object of the same name.
    

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]