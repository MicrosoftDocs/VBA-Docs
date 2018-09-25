---
title: Invalid PageTimeout setting in the Windows Registry. (Error 3249)
keywords: jeterr40.chm5003249
f1_keywords:
- jeterr40.chm5003249
ms.prod: access
ms.assetid: 62962a57-4e33-ea83-76ee-c20428deda7d
ms.date: 06/08/2017
---


# Invalid PageTimeout setting in the Windows Registry. (Error 3249)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid  **PageTimeout** setting in the Microsoft Windows Registry.

 To complete this operation


1. Exit your application.
    
2. Start the Registry Editor, and navigate to the  **PageTimeout** value. Depending on which installable ISAM you are trying to use, the invalid entry is in the **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Specify a new value in the  **Value data** box.
    
5. Restart your application, and then try the operation again.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](https://answers.microsoft.com/)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

