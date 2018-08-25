---
title: Use Default Paper Size property
ROBOTS: INDEX
keywords: vbaac10.chm5692
f1_keywords:
- vbaac10.chm5692
ms.prod: access
ms.assetid: 09d23b9f-214b-e45b-f8c2-08af26bec247
ms.date: 06/08/2017
---


# Use Default Paper Size property

**Applies to:** Access 2013 | Access 2016

You can use the **Use Default Paper Size** property to specify whether or not the default paper size of the current printer is used when you print a form or report.


## Setting

The **Use Default Paper Size** property uses the following settings.

|**Setting**|**Description**|
|:-----|:-----|
|**Yes**|Use the default paper size of the current printer when printing the form or report.|
|**No**|(Default) Use the paper size specified when the form or report was designed when printing the form or report.|

## Remarks

By default, when you print a form or report, Access uses the paper size that was specified when the form or report was designed. This can cause problems when the form or report is printed on a printer that uses a different default paper size. 

For example, you might design a form or report while connected to a printer that has a default paper size of A4. An error might occur when you print the form or report on a printer that has a default paper size of Letter. Setting the **Use Default Paper Size** property to **Yes** prevents this problem from occurring.

## See also

- [Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/en-us/msoffice/forum?page=1&;tab=question&;status=all&;auth=1)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)