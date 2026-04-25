---
title: Cannot perform this operation. (Error 3032)
keywords: jeterr40.chm5003032
f1_keywords:
- jeterr40.chm5003032
ms.assetid: 97a6b163-1ec8-176b-ee8d-d19610b29239
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Cannot perform this operation. (Error 3032)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- You tried to delete the only user account in the group Admins. The group Admins must have at least one user account. If you want to delete this account, create a new account and add it to the group Admins, or add an existing account to the group Admins, and then delete the account.
    
- You tried to put a user in a group to which the user already belongs, by appending either a **Group** object to a **User** object's **Groups** collection that already has a **Group** object of the same name or a **User** object to a **Group** object's **Users** collection that already has a **User** object of the same name.
    

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]