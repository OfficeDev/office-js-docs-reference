---
title: MessageAttachment in the manifest file (preview)
description: The MessageAttachment element configures your event-based add-in to activate on the OnMessageReadWithCustomAttachment event.
ms.date: 01/18/2024
ms.localizationpriority: medium
---

# MessageAttachment element (preview)

Configures your event-based add-in to activate on the `OnMessageReadWithCustomAttachment` event.

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

## Syntax

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnMessageReadWithCustomAttachment" FunctionName="onMessageReadWithCustomAttachmentHandler">
      <MessageAttachments>
        <MessageAttachment AttachmentExtension="xml"/>
        <MessageAttachment AttachmentExtension="json"/>
      </MessageAttachments>
    </LaunchEvent>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## Contained in

- [MessageAttachments](messageattachments.md)

## Attributes

| Attribute | Required | Description |
|:-----|:-----:|:-----|
| **AttachmentExtension** | Yes | Specifies the file extension of the attachment included in a message on which the `OnMessageReadWithCustomAttachment` event occurs. The file extension value is limited to 50 characters and must not include a period. |

## Child elements

None.

## See also

- [LaunchEvent element](launchevent.md)
- [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events)
