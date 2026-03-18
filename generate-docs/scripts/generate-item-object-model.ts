// Generates the Properties and Methods tables for outlook-item-object-model.md
// from the API Extractor JSON output files.
//
// Usage (from generate-docs/scripts/):
//   npm run build && node generate-item-object-model.js
//
// Reads from: generate-docs/json/outlook*/outlook.api.json
//             generate-docs/json/office/office.api.json
// Writes to:  docs/includes/outlook-item-object-model-properties.md
//             docs/includes/outlook-item-object-model-methods.md
//             docs/includes/outlook-item-object-model-events.md

import * as fs from "fs";
import * as path from "path";

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

/** The 4 interfaces that make up the Outlook Item object model. */
const TARGET_INTERFACES: Record<string, string> = {
    "Office.AppointmentCompose": "Appointment Organizer",
    "Office.AppointmentRead": "Appointment Attendee",
    "Office.MessageCompose": "Message Compose",
    "Office.MessageRead": "Message Read",
};

/** Map interface names to their URL-safe identifiers. */
const INTERFACE_URL_NAMES: Record<string, string> = {
    "Office.AppointmentCompose": "office.appointmentcompose",
    "Office.AppointmentRead": "office.appointmentread",
    "Office.MessageCompose": "office.messagecompose",
    "Office.MessageRead": "office.messageread",
};

/** Display order for modes within a grouped row. */
const MODE_ORDER = [
    "Office.AppointmentCompose",
    "Office.AppointmentRead",
    "Office.MessageCompose",
    "Office.MessageRead",
];

/** The query string appended to all API reference links. */
const VIEW_PARAMS = "?view=outlook-js-preview&preserve-view=true";

/** Primitive types that should NOT be linked (all lowercase for case-insensitive lookup). */
const PRIMITIVES = new Set(["string", "number", "boolean", "void", "date", "any", "undefined", "null", "object"]);

/**
 * Manual overrides for the requirement set column.
 * Use when the JSON source doesn't capture platform-specific or
 * feature-specific requirement set distinctions.
 * The value replaces the entire requirement set cell for ALL modes of that member.
 */
const REQUIREMENT_SET_OVERRIDES: Record<string, string> = {
    "addFileAttachmentAsync":
        "[1.1](outlook-requirement-set-1-1.md)<br>(classic Windows, Mac)" +
        "<br><br>[1.8](outlook-requirement-set-1-8.md)<br>(Web, new Windows)",
    "getSharedPropertiesAsync":
        "[1.8](outlook-requirement-set-1-8.md)<br>(shared folder support)" +
        "<br><br>[1.13](outlook-requirement-set-1-13.md)<br>(shared mailbox support)",
};

/** Directory for generated include files. */
const INCLUDES_DIR = path.resolve(__dirname, "../../docs/includes");

/** Directory containing JSON API model files. */
const JSON_DIR = path.resolve(__dirname, "../json");

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface ExcerptToken {
    kind: string;
    text: string;
    canonicalReference?: string;
}

interface Parameter {
    parameterName: string;
    parameterTypeTokenRange: { startIndex: number; endIndex: number };
    isOptional: boolean;
}

interface ApiMember {
    kind: string;
    name: string;
    docComment?: string;
    excerptTokens: ExcerptToken[];
    propertyTypeTokenRange?: { startIndex: number; endIndex: number };
    returnTypeTokenRange?: { startIndex: number; endIndex: number };
    parameters?: Parameter[];
}

interface ApiInterface {
    kind: string;
    name: string;
    members: ApiMember[];
}

interface MemberInfo {
    interfaceName: string;
    member: ApiMember;
    permissionLevel: string;
}

interface GroupedMember {
    name: string;
    kind: string; // "PropertySignature" | "MethodSignature"
    isDeprecated: boolean;
    entries: MemberInfo[];
    allOverloads: ApiMember[];
}

// ---------------------------------------------------------------------------
// JSON Loading
// ---------------------------------------------------------------------------

function loadApiJson(filePath: string): any {
    const raw = fs.readFileSync(filePath, "utf-8");
    return JSON.parse(raw);
}

/**
 * Auto-detects versioned Outlook JSON directories (e.g., "outlook_1_1", "outlook_1_15")
 * by scanning the JSON directory, then returns them sorted oldest to newest.
 */
function discoverOutlookVersions(): string[] {
    const entries = fs.readdirSync(JSON_DIR);
    const versionDirs = entries.filter((name: string) => /^outlook_\d+_\d+$/.test(name));
    versionDirs.sort((a: string, b: string) => {
        const [, aMajor, aMinor] = a.match(/^outlook_(\d+)_(\d+)$/)!;
        const [, bMajor, bMinor] = b.match(/^outlook_(\d+)_(\d+)$/)!;
        return (+aMajor - +bMajor) || (+aMinor - +bMinor);
    });
    return versionDirs;
}

/**
 * Navigates Package → EntryPoint → Namespace("Office") → finds target interfaces.
 */
function extractInterfaces(apiJson: any): Map<string, ApiInterface> {
    const result = new Map<string, ApiInterface>();
    const entryPoint = apiJson.members?.[0];
    if (!entryPoint) return result;

    // Find the "Office" namespace within the entry point
    for (const member of entryPoint.members || []) {
        if (member.kind === "Namespace" && member.name === "Office") {
            for (const iface of member.members || []) {
                const qualifiedName = `Office.${iface.name}`;
                if (iface.kind === "Interface" && TARGET_INTERFACES[qualifiedName]) {
                    result.set(qualifiedName, iface);
                }
            }
        }
    }
    return result;
}

// ---------------------------------------------------------------------------
// Metadata extraction from docComment
// ---------------------------------------------------------------------------

function extractPermissionLevel(docComment: string): string {
    // Pattern: **{@link ... | Minimum permission level}**: **read item**
    const match = docComment.match(/Minimum permission level\}[*]*:\s*\*\*([^*]+)\*\*/);
    if (match) return match[1].trim();
    return "";
}

function extractApiSetVersion(docComment: string): string {
    // Pattern: [Api set: Mailbox X.Y ...]
    const match = docComment.match(/\[Api set:\s*Mailbox\s+([\d.]+)/i);
    if (match) return match[1];
    return "";
}

function isDeprecated(docComment: string): boolean {
    return /@deprecated/i.test(docComment);
}

// ---------------------------------------------------------------------------
// Type formatting
// ---------------------------------------------------------------------------

/**
 * Converts a canonical reference to a documentation URL path.
 * e.g., "outlook!Office.Body:interface" → "/javascript/api/outlook/office.body"
 */
function canonicalRefToUrl(ref: string): string {
    // Format: "package!Namespace.Type:kind" or "package!Namespace.Type.SubType:kind"
    const match = ref.match(/^([^!]+)!(.+):(\w+)$/);
    if (!match) return "";
    const pkg = match[1];
    const typePath = match[2].toLowerCase();
    return `/javascript/api/${pkg}/${typePath}`;
}

/**
 * Formats a return/property type from excerpt tokens into markdown.
 */
function formatType(tokens: ExcerptToken[], range: { startIndex: number; endIndex: number }): string {
    const parts: string[] = [];
    for (let i = range.startIndex; i < range.endIndex; i++) {
        const token = tokens[i];
        if (token.kind === "Reference" && token.canonicalReference) {
            const displayName = token.text;
            if (PRIMITIVES.has(displayName.toLowerCase())) {
                parts.push(capitalizeType(displayName));
            } else {
                const url = canonicalRefToUrl(token.canonicalReference);
                if (url) {
                    parts.push(`[${displayName}](${url}${VIEW_PARAMS})`);
                } else {
                    parts.push(displayName);
                }
            }
        } else {
            // Content token: could be "[]", ", ", primitives like "string", etc.
            // Capitalize standalone primitive types.
            const trimmed = token.text.trim();
            if (PRIMITIVES.has(trimmed)) {
                parts.push(token.text.replace(trimmed, capitalizeType(trimmed)));
            } else if (/\|\s*string\b/.test(token.text)) {
                // Union types like " | string" — drop the primitive union portion for cleaner display.
                parts.push(token.text.replace(/\s*\|\s*string\b/, ""));
            } else {
                parts.push(token.text);
            }
        }
    }

    let result = parts.join("");

    // Convert TypeScript array notation "Type[]" to "Array.<Type>" for consistency
    // with the existing table format.
    if (result.endsWith("[]")) {
        const inner = result.slice(0, -2);
        result = `Array.<${inner}>`;
    }

    return result;
}

/** Capitalize primitive type names to match existing table style. */
function capitalizeType(t: string): string {
    const map: Record<string, string> = {
        string: "String",
        number: "Number",
        boolean: "Boolean",
        void: "void",
        date: "Date",
    };
    return map[t.toLowerCase()] || t;
}

// ---------------------------------------------------------------------------
// Method signature formatting
// ---------------------------------------------------------------------------

/**
 * Builds a display signature like: methodName(param1, param2, [optionalParam])
 * Merges all overloads across interfaces: picks the overload with the most
 * parameters, then marks a parameter as optional if it's explicitly optional
 * OR if any overload omits that parameter name.
 */
function formatMethodSignature(memberEntries: MemberInfo[], allOverloads: ApiMember[]): string {
    // Pick the overload with the most parameters.
    let bestOverload: ApiMember | null = null;
    let maxParams = -1;
    for (const m of allOverloads) {
        const count = m.parameters?.length || 0;
        if (count > maxParams) {
            maxParams = count;
            bestOverload = m;
        }
    }

    if (!bestOverload) return memberEntries[0].member.name + "()";

    // Collect all parameter names present in each overload.
    const overloadParamSets = allOverloads.map(
        (m) => new Set((m.parameters || []).map((p) => p.parameterName))
    );

    const params = bestOverload.parameters || [];
    const paramParts = params.map((p) => {
        // A parameter is optional if it's marked optional OR if any overload
        // doesn't include it.
        const missingInSomeOverload = overloadParamSets.some(
            (paramSet) => !paramSet.has(p.parameterName)
        );
        const isOpt = p.isOptional || missingInSomeOverload;
        return isOpt ? `[${p.parameterName}]` : p.parameterName;
    });

    return `${bestOverload.name}(${paramParts.join(", ")})`;
}

// ---------------------------------------------------------------------------
// Minimum requirement set detection
// ---------------------------------------------------------------------------

/**
 * For each versioned JSON, checks whether a member exists in the given interface.
 * Returns the earliest version where it appears, or "Preview" if only in the
 * base (preview) file.
 */
function findMinimumRequirementSet(
    memberName: string,
    interfaceName: string,
    memberKind: string,
    versionedInterfaces: Map<string, Map<string, ApiInterface>>,
    outlookVersions: string[]
): string {
    for (const version of outlookVersions) {
        const interfaces = versionedInterfaces.get(version);
        if (!interfaces) continue;

        const iface = interfaces.get(interfaceName);
        if (!iface) continue;

        const found = iface.members.some(
            (m) => m.name === memberName && m.kind === memberKind
        );
        if (found) {
            // Convert "outlook_1_8" → "1.8"
            const verNum = version.replace("outlook_", "").replace("_", ".");
            return verNum;
        }
    }

    return "Preview";
}

// ---------------------------------------------------------------------------
// Markdown generation
// ---------------------------------------------------------------------------

function buildModeLink(interfaceName: string, memberName: string, memberKind: string): string {
    const displayName = TARGET_INTERFACES[interfaceName];
    const urlIface = INTERFACE_URL_NAMES[interfaceName];
    const memberLower = memberName.toLowerCase();
    const anchor = memberKind === "MethodSignature"
        ? `#outlook-office-${urlIface.replace("office.", "")}-${memberLower}-member(1)`
        : `#outlook-office-${urlIface.replace("office.", "")}-${memberLower}-member`;

    return `[${displayName}](/javascript/api/outlook/${urlIface}${VIEW_PARAMS}${anchor})`;
}

function buildVersionLink(version: string): string {
    if (version === "Preview") {
        return `[Preview](outlook-requirement-set-preview.md)`;
    }
    // "1.8" → "outlook-requirement-set-1-8.md"
    const fileSuffix = version.replace(".", "-");
    return `[${version}](outlook-requirement-set-${fileSuffix}.md)`;
}

function generatePropertiesTable(groups: GroupedMember[], versionedInterfaces: Map<string, Map<string, ApiInterface>>, outlookVersions: string[]): string {
    const lines: string[] = [];
    lines.push("| Property | Minimum permission level | Details by mode | Return type | Minimum requirement set |");
    lines.push("| --- | --- | --- | --- | :---: |");

    for (const group of groups) {
        if (group.kind !== "PropertySignature") continue;

        const sortedEntries = sortByModeOrder(group.entries);
        let isFirst = true;
        let lastPermission = "";

        for (const entry of sortedEntries) {
            const modeLink = buildModeLink(entry.interfaceName, group.name, group.kind);
            const typeRange = entry.member.propertyTypeTokenRange;
            const returnType = typeRange
                ? formatType(entry.member.excerptTokens, typeRange)
                : "";
            const versionLink = REQUIREMENT_SET_OVERRIDES[group.name]
                ?? buildVersionLink(findMinimumRequirementSet(
                    group.name, entry.interfaceName, group.kind, versionedInterfaces, outlookVersions
                ));

            if (isFirst) {
                const nameCol = group.isDeprecated ? `${group.name} **(deprecated)**` : group.name;
                lines.push(`| ${nameCol} | **${entry.permissionLevel}** | ${modeLink} | ${returnType} | ${versionLink} |`);
                lastPermission = entry.permissionLevel;
                isFirst = false;
            } else {
                const permCol = entry.permissionLevel !== lastPermission
                    ? ` **${entry.permissionLevel}**`
                    : "";
                lines.push(`| |${permCol} | ${modeLink} | ${returnType} | ${versionLink} |`);
                lastPermission = entry.permissionLevel;
            }
        }
    }

    return lines.join("\n");
}

function generateMethodsTable(groups: GroupedMember[], versionedInterfaces: Map<string, Map<string, ApiInterface>>, outlookVersions: string[]): string {
    const lines: string[] = [];
    lines.push("| Method | Minimum permission level | Details by mode | Minimum requirement set |");
    lines.push("| --- | --- | --- | :---: |");

    for (const group of groups) {
        if (group.kind !== "MethodSignature") continue;

        const sortedEntries = sortByModeOrder(group.entries);
        const signature = formatMethodSignature(group.entries, group.allOverloads);
        let isFirst = true;
        let lastPermission = "";

        for (const entry of sortedEntries) {
            const modeLink = buildModeLink(entry.interfaceName, group.name, group.kind);
            const versionLink = REQUIREMENT_SET_OVERRIDES[group.name]
                ?? buildVersionLink(findMinimumRequirementSet(
                    group.name, entry.interfaceName, group.kind, versionedInterfaces, outlookVersions
                ));

            if (isFirst) {
                const nameCol = group.isDeprecated ? `${signature} **(deprecated)**` : signature;
                lines.push(`| ${nameCol} | **${entry.permissionLevel}** | ${modeLink} | ${versionLink} |`);
                lastPermission = entry.permissionLevel;
                isFirst = false;
            } else {
                const permCol = entry.permissionLevel !== lastPermission
                    ? ` **${entry.permissionLevel}**`
                    : "";
                lines.push(`| |${permCol} | ${modeLink} | ${versionLink} |`);
                lastPermission = entry.permissionLevel;
            }
        }
    }

    return lines.join("\n");
}

function generateEventsTable(): string {
    // Load the Office JSON (EventType is in the Office namespace, not Outlook).
    const officeJsonPath = path.join(JSON_DIR, "office", "office.api.json");
    const officeJson = loadApiJson(officeJsonPath);
    // Navigate: Package → EntryPoint → Namespace("Office") → Enum("EventType")
    const entryPoint = officeJson.members[0];
    const officeNs = entryPoint.members.find((m: any) => m.name === "Office");
    const eventType = officeNs.members.find((m: any) => m.name === "EventType");

    // Auto-detect item-relevant events: has Mailbox API set AND either
    // mentions the Item object or mentions neither Item nor Mailbox object
    // (e.g., SpamReporting which is event-based activation, not addHandlerAsync).
    const itemEvents = eventType.members.filter((m: any) => {
        const doc = m.docComment || "";
        const hasMailboxApiSet = /\[Api set:\s*Mailbox/i.test(doc);
        if (!hasMailboxApiSet) return false;
        const mentionsItemObj = /method of the .Item. object/i.test(doc);
        const mentionsMailboxObj = /method of the .Mailbox. object/i.test(doc);
        return mentionsItemObj || !mentionsMailboxObj;
    });

    const lines: string[] = [];
    lines.push("| Event | Description | Minimum requirement set |");
    lines.push("| --- | --- | :---: |");

    for (const member of itemEvents) {
        const docComment = member.docComment || "";
        const description = extractEventDescription(docComment);
        const version = extractApiSetVersion(docComment);
        const versionLink = buildVersionLink(version || "Preview");
        lines.push(`| \`${member.name}\` | ${description} | ${versionLink} |`);
    }

    console.log(`Found ${itemEvents.length} item events.`);
    return lines.join("\n");
}

/**
 * Extracts a description from an EventType member's docComment.
 * Pulls the first sentence, cleans up {@link} markup, and appends
 * a task pane or function command availability note.
 */
function extractEventDescription(docComment: string): string {
    // Strip the leading "/**\n * " and trailing " */\n".
    const cleaned = docComment
        .replace(/^\/\*\*\s*\n/m, "")
        .replace(/\s*\*\/\s*$/m, "")
        .replace(/^ \* ?/gm, "")
        .trim();

    // Replace {@link url | display text} with just the display text.
    // Also handle {@link url} (no display text) by stripping the whole thing.
    let text = cleaned
        .replace(/\{@link\s+[^|}]+\|\s*([^}]+)\}/g, "$1")
        .replace(/\{@link\s+[^}]+\}/g, "")
        .trim();

    // Extract the first sentence (up to the first period followed by whitespace or end).
    const firstSentenceMatch = text.match(/^(.+?\.)(\s|$)/);
    let description = firstSentenceMatch ? firstSentenceMatch[1].trim() : text;

    // Determine availability from the full docComment.
    if (/task pane.*function commands can't/i.test(docComment) ||
        /can only be handled in a task pane/i.test(docComment)) {
        description += " Only available with task pane implementation.";
    }

    return description;
}

function sortByModeOrder(entries: MemberInfo[]): MemberInfo[] {
    return [...entries].sort(
        (a, b) => MODE_ORDER.indexOf(a.interfaceName) - MODE_ORDER.indexOf(b.interfaceName)
    );
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

function main(): void {
    console.log("Loading preview API JSON...");
    const previewJsonPath = path.join(JSON_DIR, "outlook", "outlook.api.json");
    const previewJson = loadApiJson(previewJsonPath);
    const previewInterfaces = extractInterfaces(previewJson);

    console.log(`Found ${previewInterfaces.size} target interfaces in preview JSON.`);

    // Load all versioned JSONs for minimum requirement set detection.
    console.log("Loading versioned API JSONs...");
    const outlookVersions = discoverOutlookVersions();
    const versionedInterfaces = new Map<string, Map<string, ApiInterface>>();
    for (const version of outlookVersions) {
        const versionPath = path.join(JSON_DIR, version, "outlook.api.json");
        if (fs.existsSync(versionPath)) {
            const versionJson = loadApiJson(versionPath);
            versionedInterfaces.set(version, extractInterfaces(versionJson));
        }
    }
    console.log(`Loaded ${versionedInterfaces.size} versioned JSONs.`);

    // Collect all members across the 4 interfaces, grouped by member name.
    const memberGroups = new Map<string, GroupedMember>();

    for (const [interfaceName, iface] of previewInterfaces) {
        // Track seen method names to only take overload index 1 for table entries.
        const seenMethods = new Set<string>();

        for (const member of iface.members) {
            if (member.kind !== "PropertySignature" && member.kind !== "MethodSignature") {
                continue;
            }

            const key = `${member.name}|${member.kind}`;

            // Ensure the group exists.
            if (!memberGroups.has(key)) {
                memberGroups.set(key, {
                    name: member.name,
                    kind: member.kind,
                    isDeprecated: false,
                    entries: [],
                    allOverloads: [],
                });
            }

            const group = memberGroups.get(key)!;

            // Collect all overloads for method signature merging.
            if (member.kind === "MethodSignature") {
                group.allOverloads.push(member);
            }

            // For methods, only add one table entry per interface (first overload).
            if (member.kind === "MethodSignature" && seenMethods.has(member.name)) {
                continue;
            }
            if (member.kind === "MethodSignature") {
                seenMethods.add(member.name);
            }

            const docComment = member.docComment || "";
            const permissionLevel = extractPermissionLevel(docComment);
            const deprecated = isDeprecated(docComment);

            // If any interface marks it deprecated, the whole group is deprecated.
            if (deprecated) {
                group.isDeprecated = true;
            }

            group.entries.push({
                interfaceName,
                member,
                permissionLevel,
            });
        }
    }

    // Sort groups alphabetically by name.
    const sortedGroups = [...memberGroups.values()].sort((a, b) =>
        a.name.localeCompare(b.name)
    );

    console.log(`Found ${sortedGroups.filter(g => g.kind === "PropertySignature").length} properties and ${sortedGroups.filter(g => g.kind === "MethodSignature").length} methods.`);

    // Generate tables.
    const propertiesTable = generatePropertiesTable(sortedGroups, versionedInterfaces, outlookVersions);
    const methodsTable = generateMethodsTable(sortedGroups, versionedInterfaces, outlookVersions);

    // Write include files.
    const propsFile = path.join(INCLUDES_DIR, "outlook-item-object-model-properties.md");
    fs.writeFileSync(propsFile, propertiesTable + "\n", "utf-8");
    console.log(`Wrote ${propsFile}`);

    const methodsFile = path.join(INCLUDES_DIR, "outlook-item-object-model-methods.md");
    fs.writeFileSync(methodsFile, methodsTable + "\n", "utf-8");
    console.log(`Wrote ${methodsFile}`);

    // Generate Events table from Office.EventType enum.
    const eventsTable = generateEventsTable();
    const eventsFile = path.join(INCLUDES_DIR, "outlook-item-object-model-events.md");
    fs.writeFileSync(eventsFile, eventsTable + "\n", "utf-8");
    console.log(`Wrote ${eventsFile}`);
}

main();
