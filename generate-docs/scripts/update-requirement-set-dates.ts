#!/usr/bin/env node --harmony

import * as fsx from 'fs-extra';
import * as path from "path";
import * as crypto from "crypto";

/**
 * Updates `ms.date` in requirement-set pages when their included API table files
 * have changed since the last doc generation run.
 *
 * Uses a content-hash manifest (include-hashes.json) to detect changes. On each
 * run, hashes all generated include files, compares to the stored manifest,
 * updates ms.date for pages whose includes changed, then writes the updated
 * manifest.
 *
 * This script should run after all include files in docs/includes/ have been
 * generated (i.e., after the midprocessor and API Documenter steps).
 */

const REPO_ROOT = path.resolve(__dirname, "../..");
const REQUIREMENT_SETS_DIR = path.resolve(REPO_ROOT, "docs/requirement-sets");
const INCLUDES_DIR = path.resolve(REPO_ROOT, "docs/includes");
const HASH_MANIFEST_PATH = path.resolve(__dirname, "../script-inputs/include-hashes.json");

// Only track generated API-table includes (not shared/static includes).
const GENERATED_INCLUDE_PATTERN = /^(excel|word|powerpoint|outlook|onenote|visio)[-_].*\.md$/;

type HashManifest = Record<string, string>;

function computeFileHash(filePath: string): string {
    const content = fsx.readFileSync(filePath);
    return crypto.createHash("sha256").update(content).digest("hex");
}

function loadManifest(): HashManifest {
    if (fsx.existsSync(HASH_MANIFEST_PATH)) {
        return JSON.parse(fsx.readFileSync(HASH_MANIFEST_PATH, "utf8"));
    }
    return {};
}

function saveManifest(manifest: HashManifest): void {
    fsx.ensureDirSync(path.dirname(HASH_MANIFEST_PATH));
    fsx.writeFileSync(HASH_MANIFEST_PATH, JSON.stringify(manifest, null, 2) + "\n");
}

function toRelativeKey(filePath: string): string {
    return path.relative(REPO_ROOT, filePath).replace(/\\/g, "/");
}

function getGeneratedIncludes(): string[] {
    if (!fsx.existsSync(INCLUDES_DIR)) return [];
    return fsx.readdirSync(INCLUDES_DIR)
        .filter(name => GENERATED_INCLUDE_PATTERN.test(name))
        .map(name => path.join(INCLUDES_DIR, name));
}

function findChangedIncludes(oldManifest: HashManifest, newManifest: HashManifest): Set<string> {
    const changed = new Set<string>();
    for (const [relKey, newHash] of Object.entries(newManifest)) {
        if (oldManifest[relKey] !== newHash) {
            // Resolve back to absolute path for matching against include references
            changed.add(path.resolve(REPO_ROOT, relKey));
        }
    }
    return changed;
}

function getTodayFormatted(): string {
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const year = now.getFullYear();
    return `${month}/${day}/${year}`;
}

function extractMsDate(content: string): string | null {
    const match = content.match(/^ms\.date:\s*(\d{1,2}\/\d{1,2}\/\d{4})\s*$/m);
    return match ? match[1] : null;
}

function extractIncludePaths(content: string, fileDir: string): string[] {
    const regex = /\[!INCLUDE[^\]]*\]\(([^)]+\.md)\)/gi;
    const paths: string[] = [];
    let match: RegExpExecArray | null;
    while ((match = regex.exec(content)) !== null) {
        const includePath = match[1];
        const resolved = path.resolve(fileDir, includePath);
        if (resolved.startsWith(INCLUDES_DIR)) {
            const basename = path.basename(resolved);
            if (GENERATED_INCLUDE_PATTERN.test(basename)) {
                paths.push(resolved);
            }
        }
    }
    return paths;
}

function processRequirementSetPages(changedIncludes: Set<string>, today: string): number {
    let updatedCount = 0;

    function walkDir(dir: string): void {
        if (!fsx.existsSync(dir)) return;

        for (const entry of fsx.readdirSync(dir, { withFileTypes: true })) {
            const fullPath = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                walkDir(fullPath);
            } else if (entry.isFile() && entry.name.endsWith(".md")) {
                const content = fsx.readFileSync(fullPath, "utf8");
                const msDate = extractMsDate(content);
                if (!msDate) continue;

                // Skip if already set to today
                if (msDate === today) continue;

                const includePaths = extractIncludePaths(content, path.dirname(fullPath));
                const hasChangedInclude = includePaths.some(p => changedIncludes.has(p));

                if (hasChangedInclude) {
                    const updatedContent = content.replace(
                        /^ms\.date:\s*\d{1,2}\/\d{1,2}\/\d{4}\s*$/m,
                        `ms.date: ${today}`
                    );
                    fsx.writeFileSync(fullPath, updatedContent);
                    updatedCount++;
                    console.log(`  Updated: ${path.relative(REPO_ROOT, fullPath)} -> ${today}`);
                }
            }
        }
    }

    walkDir(REQUIREMENT_SETS_DIR);
    return updatedCount;
}

// Main
console.log("\nUpdating ms.date in requirement-set pages...");

const oldManifest = loadManifest();

// Build new manifest by hashing all generated include files (using repo-relative keys)
const newManifest: HashManifest = {};
for (const filePath of getGeneratedIncludes()) {
    newManifest[toRelativeKey(filePath)] = computeFileHash(filePath);
}

const changedIncludes = findChangedIncludes(oldManifest, newManifest);

if (changedIncludes.size === 0) {
    console.log("  No include file changes detected. Nothing to update.");
} else {
    console.log(`  Found ${changedIncludes.size} changed include file(s).`);
    const today = getTodayFormatted();
    const count = processRequirementSetPages(changedIncludes, today);
    console.log(`  Updated ${count} requirement-set page(s).`);
}

// Always save the updated manifest so next run has a baseline
saveManifest(newManifest);
console.log("Done.\n");
