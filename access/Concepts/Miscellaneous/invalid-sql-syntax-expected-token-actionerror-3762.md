---
title: "Invalid SQL Syntax: expected token: ACTION. (Error 3762)"
ms.prod: access
ms.assetid: 73122947-9db6-f417-7e34-96bc4108bab3
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid SQL Syntax: expected token: ACTION. (Error 3762)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when defining referential integrity constraints through the CREATE TABLE syntax or the ALTER TABLE ALTER COLUMN syntax. It occurs when the keyword NO is not followed by the keyword ACTION. For example, by omitting the BOLD ON keyword, the following would generate the error:

CREATE TABLE OrderDetail (OrderId LONG CONSTRAINT fkOrdersOrderId REFERENCES Orders (OrderId) ON UPDATE CASCADE ON DELETE  **NO** ACTION, LineItem LONG, ProductID LONG CONSTRAINT fkProductsProductId REFERENCES Products (ProductId), Quantity LONG);

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]