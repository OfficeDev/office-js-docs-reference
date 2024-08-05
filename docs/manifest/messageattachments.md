---
title: MessageAttachments in the manifest file (preview)
description: The MessageAttachments element configures your event-based add-in to activate on the OnMessageReadWithCustomAttachment event.
ms.date: 01/18/2024
ms.localizationpriority: medium
---

# MessageAttachments element (preview)

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

- [LaunchEvent](launchevent.md) (when the `Type` attribute is set to `OnMessageReadWithCustomAttachment`)

## Attributes

None.

## Child elements

| Element | Required | Description |
|:-----|:-----:|:-----|
| [MessageAttachment](messageattachment.md) | Yes | Configures your add-in to activate on the `OnMessageReadWithCustomAttachment` event. You can specify a maximum of two **\<MessageAttachment\>** elements in your manifest. |

## See also

- [LaunchEvent element](launchevent.md)
- [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch#supported-events)
