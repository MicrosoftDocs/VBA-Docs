---
title: Support for keeping add-ins enabled
ms.assetid: 63cd5a19-6c46-42f9-8fe2-9ce943bf106c
ms.date: 11/12/2018
ms.localizationpriority: medium
---


# Support for keeping add-ins enabled

Programs in Office 2013 and later versions provide add-in resiliency, meaning that apps will disable an add-in if it performs slowly. However, you can re-enable add-ins and prevent add-ins from being auto-disabled by other Office programs. 

## Preventing add-ins from being disabled

While most add-ins will not be disabled by the add-in disabling feature, you don't want your add-in to be disabled consistently.

Following are some suggestions for improving add-in performance:

- Prefer native COM add-ins over managed add-ins because managed add-ins must incur the overhead of loading the .NET Framework during Outlook startup.
    
- If you have long-running tasks such as making an expensive connection to a database, defer those tasks to occur after startup.
    
- If possible, cache data locally rather than making expensive network calls during the **FolderSwitch** and **BeforeFolderSwitch** events of an explorer, or **Open** events of an item.
    
- Be aware that all calls to the Outlook object model execute on Outlook's main foreground thread. Avoid making long-running Outlook object model calls if possible.

- In Outlook 2013, calls to the Outlook object model return E_RPC_WRONG_THREAD when the Outlook object model is called from a background thread.
 
- Polling is an expensive operation, so always prefer an event-driven model over polling.

> [!NOTE]
> You cannot prevent Outlook from disabling add-ins in the following conditions:
> 
> - The add-in crashes Outlook.
> 
> - The add-in cannot be loaded.
> 
> In these cases, the cause of the crash or the loading failure needs to be fixed together with the add-in owner.

## System administrator control over add-ins

The user has control over which add-ins run on their computer. Beginning with Office 2013, system administrators can configure an enhanced level of control for add-ins by using group policy. Group policy will always override user settings and users are prevented from changing add-in settings for add-ins that have been configured by the group policy **List of Managed Add-ins**.

For Outlook, the registry keys and settings are described in the following tables.

|Name|Description|
|:-----|:-----|
|Key|Office 2013:<br />HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Resiliency\AddinList<br /><br />Office 2016/2019/365:<br />HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList|
|Description|This policy setting allows you to specify the list of managed add-ins that are always enabled, always disabled (blocked), or configurable by the user.<br/><br/>**NOTE**: Here, the term "managed" refers to add-ins that are handled by the group policy, and does not relate to add-ins being developed in managed programming languages.|
|**String**|ProgID of the add-in|
|Values|Specify the value as follows:<br>0 = always disabled (blocked)<br>1 = always enabled<br>2 = configurable by the user and not blocked by the **Block all unmanaged add-ins** policy setting when enabled.|

> [!NOTE]
> To obtain the ProgID for an add-in, use the Windows Registry Editor on the client computer where the add-in is installed. Copy the registry key names under found: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\\\<application\>\Addins or HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\\\<application\>\Addins. Registry key names are **case sensitive**.
> 
> - If you disable or don't enable this policy setting, the list of managed add-ins will be deleted. If the **Block all unmanaged add-ins** policy setting is enabled, then all add-ins are blocked.
> 
> - Add-ins that are disabled by this policy will never be disabled by the Outlook add-in disabling feature, which disables add-ins for performance, resiliency, or reliability reasons.
> 
> - If the user chooses "Always enable this add-in", the registry is updated to include details about the add-in that is to be exempted from the automatic disabling feature.

|Name|Description|
|:-----|:-----|
|Key|HKEY_CURRENT_USER\Software\Microsoft\Office\x.0\Outlook\Resiliency\DoNotDisableAddinList|
|Description|This Key prevents add-ins you specify from being disabled by the add-in disabling feature.|
|**DWORD**|ProgID of the add-in|
|Values| A Hex value between 1 and A indicating the reason the add-in was originally disabled:<br>0x00000001 Boot load (LoadBehavior = 3)<br>0x00000002 Demand load (LoadBehavior = 9)<br>0x00000003 Crash<br>0x00000004 Handling FolderSwitch event<br>0x00000005 Handling BeforeFolderSwitch event<br>0x00000006 Item Open<br>0x00000007 Iteration Count<br>0x00000008 Shutdown<br>0x00000009 Crash, but not disabled because add-in is in the allow list<br>0x0000000A Crash, but not disabled because user selected no in disable dialog<br/><br/>**NOTE**: The x.0 placeholder represents the version of Office (16.0 = Office 2016/2019/365, 15.0 = Office 2013).|

> [!NOTE]
> If you re-enable an add-in that caused a performance problem at one time, users may experience performance problems in the future in the Office program for which the add-in is loaded.

To block add-ins that are not managed by this policy setting, you must also configure the **Block all unmanaged add-ins** policy setting.

## See also

- [Concepts (Outlook)](concepts-outlook-vba-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
