---
title: "Invalid SQL Syntax: expected token: ACTION. (Error 3762)"
ms.assetid: 73122947-9db6-f417-7e34-96bc4108bab3
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Invalid SQL Syntax: expected token: ACTION. (Error 3762)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when defining referential integrity constraints through the CREATE TABLE syntax or the ALTER TABLE ALTER COLUMN syntax. It occurs when the keyword NO is not followed by the keyword ACTION. For example, by omitting the BOLD ON keyword, the following would generate the error:

CREATE TABLE OrderDetail (OrderId LONG CONSTRAINT fkOrdersOrderId REFERENCES Orders (OrderId) ON UPDATE CASCADE ON DELETE **NO** ACTION, LineItem LONG, ProductID LONG CONSTRAINT fkProductsProductId REFERENCES Products (ProductId), Quantity LONG);

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]