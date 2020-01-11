---
title: Register a custom business object
ROBOTS: INDEX
ms.prod: access
ms.assetid: eed3b78e-310a-98fa-5cf9-32edaab0402f
ms.date: 06/08/2017
localization_priority: Normal
---


# Register a custom business object

**Applies to:** Access 2013 | Access 2016

To successfully launch a custom business object (.dll or .exe) through the Web server, the business object's ProgID must be entered into the registry as explained in this procedure. This RDS feature protects the security of your Web server by running only sanctioned executables.

> [!NOTE] 
> For MDAC 2.0 and later, the default business object **RDSServer.DataFactory** is not registered by default during MDAC installation. However, if **RDSServer.DataFactory** was registered as safe for execution on the computer prior to the installation, the registry entry is maintained for the new installation.

**To register a custom business object**

1. Click **Start**, and then click **Run**.
    
2. Type **RegEdit**, and then click **OK**.
    
3. In the Registry Editor, navigate to the **HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\W3SVC\Parameters\ADCLaunch** registry key.
    
4. Select the **ADCLaunch** key, and then from the **Edit** menu, point to **New** and click **Key**.
    
5. Type the ProgID of your custom business object, and press **Enter**. Leave the **Value** entry blank.
    
## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]