import {
    App, Plugin, Workspace, WorkspaceLeaf, WorkspaceRoot, WorkspaceFloating, View, TFile, PaneType, WorkspaceTabs,
    WorkspaceItem, Platform, Keymap, Notice,
} from 'obsidian';
import * as monkeyAround from 'monkey-around';
import {
    OpenTabSettingsPluginSettingTab, OpenTabSettingsPluginSettings, DEFAULT_SETTINGS, NEW_TAB_TAB_GROUP_PLACEMENTS,
} from './settings';
import { TabGroup } from './types';


/**
 * Special view types added by plugins that should be deduplicated like normal files.
 * This is only needed if the view is not registered as the default view for a file extension.
 */
const PLUGIN_VIEW_TYPES: Record<string, string[]> = {
    "md": ["excalidraw", "kanban"],
}


function isEmptyLeaf(leaf: WorkspaceLeaf) {
    // home-tab plugin replaces new tab with home tabs, which should be treated like empty.
    return ["empty", "home-tab-view"].includes(leaf.view.getViewType())
}

/** Check if leaf is in the main area (e.g. not in sidebar etc) */
function isMainLeaf(leaf: WorkspaceLeaf) {
    const root = leaf.getRoot();
    return (root instanceof WorkspaceRoot || root instanceof WorkspaceFloating);
}

function capitalize(s: string) {
    return s[0].toUpperCase() + s.slice(1);
}

/**
 * This is a bit hacky, but to support easily changing our settings in Mod click or menu items we're sticking the
 * overrides onto the string passed to getLeaf.
 */
function buildOverride(mode: PaneType|false, settings: Partial<OpenTabSettingsPluginSettings>) {
    return `${mode || ""}:${JSON.stringify(settings)}` as PaneType; // Deceptive cast to allow passing to getLeaf
}

function parseOverride(override?: string|boolean): [PaneType|false, Partial<OpenTabSettingsPluginSettings>] {
    if (!override) {
        return [false, {}];
    } else if (override === true) {
        return ['tab', {}];
    } else {
        const [mode, ...rest] = override.split(":");
        const json = rest.join(":") || "{}";
        return [(mode || false) as PaneType|false, JSON.parse(json)];
    }
}

const OVERRIDES = {
    tab: "tab",
    same: buildOverride(false, {openInNewTab: false}),
    allow_duplicate: buildOverride(false, {deduplicateTabs: false}),
    opposite: buildOverride("tab", {newTabTabGroupPlacement: "opposite"}),
}


export default class OpenTabSettingsPlugin extends Plugin {
    settings: OpenTabSettingsPluginSettings = {...DEFAULT_SETTINGS};
    private currentPreviewLeafId: string | null = null;
    private nextOpenIsPreview: boolean = false;
    private previewHandlersRegistered: boolean = false;
    private boundClickHandler: ((evt: MouseEvent) => void) | null = null;
    private boundDblClickHandler: ((evt: MouseEvent) => void) | null = null;

    async onload() {
        await this.loadSettings();

        this.addSettingTab(new OpenTabSettingsPluginSettingTab(this.app, this));

        this.registerMonkeyPatches();
        this.registerPreviewTabHandlers();

        this.registerEvent(
            this.app.workspace.on("file-menu", (menu, file, source, leaf) => {
                if (file instanceof TFile) {
                    if (this.settings.openInNewTab) {
                        menu.addItem((item) => {
                            item.setSection("open");
                            item.setIcon("file-minus")
                            item.setTitle("Open in same tab");
                            item.onClick(async () => {
                                await this.app.workspace.getLeaf(OVERRIDES.same).openFile(file);
                            });
                        });
                    }
                    if (this.settings.deduplicateTabs && this.findMatchingLeaves(file).length > 0) {
                        menu.addItem((item) => {
                            item.setSection("open");
                            item.setIcon("files")
                            item.setTitle("Open in duplicate tab");
                            item.onClick(async () => {
                                await this.app.workspace.getLeaf(OVERRIDES.allow_duplicate).openFile(file);
                            });
                        });
                    }
                    const activeLeaf = this.app.workspace.getMostRecentLeaf();
                    if (activeLeaf && this.getAllTabGroups(activeLeaf.getRoot()).length > 1) {
                        menu.addItem((item) => {
                            item.setSection("open");
                            item.setIcon("lucide-split-square-horizontal")
                            item.setTitle("Open in opposite tab group");
                            item.onClick(async () => {
                                await this.app.workspace.getLeaf(OVERRIDES.opposite).openFile(file);
                            });
                        });
                    }
                }
            })
        );

        const commands = [
            ["openInNewTab", "always open in new tab"],
            ["deduplicateTabs", "prevent duplicate tabs"],
        ] as const;
        for (const [setting, name] of commands) {
            const id = setting.replace(/[A-Z]/g, l => `-${l.toLowerCase()}`);

            this.addCommand({
                id: `toggle-${id}`, name: `Toggle ${name}`,
                callback: async () => {
                    await this.updateSettings({[setting]: !this.settings[setting]});
                    new Notice(`${capitalize(name)} ${this.settings[setting] ? 'ON' : 'OFF'}`, 2500);
                },
            });
            this.addCommand({
                id: `enable-${id}`, name: `Enable ${name}`,
                callback: async () => {
                    await this.updateSettings({[setting]: true});
                    new Notice(`${capitalize(name)} ${this.settings[setting] ? 'ON' : 'OFF'}`, 2500);
                },
            });
            this.addCommand({
                id: `disable-${id}`, name: `Disable ${name}`,
                callback: async () => {
                    await this.updateSettings({[setting]: false});
                    new Notice(`${capitalize(name)} ${this.settings[setting] ? 'ON' : 'OFF'}`, 2500);
                },
            });
        }
        this.addCommand({
            id: `cycle-tab-group-placement`, name: `Cycle tab group placement`,
            callback: async () => {
                const values = Object.keys(NEW_TAB_TAB_GROUP_PLACEMENTS) as (keyof typeof NEW_TAB_TAB_GROUP_PLACEMENTS)[];
                const index = values.findIndex(v => v == this.settings.newTabTabGroupPlacement);
                const newValue = values[(index + 1) % values.length];
                await this.updateSettings({newTabTabGroupPlacement: newValue});
                new Notice(`Tab group placement: ${NEW_TAB_TAB_GROUP_PLACEMENTS[newValue]}`, 2500);
            },
        });
    }

    registerMonkeyPatches() {
        const registerPatches = (plugin: OpenTabSettingsPlugin) => {
            plugin.register(monkeyAround.around(Workspace.prototype, {
                /**
                 * Patch getLeaf to open leaves in new tab by default, based on settings.
                 */
                getLeaf: (oldMethod) => {
                    return function(this: Workspace, openModeIn?: string|boolean, ...args: unknown[]) {
                        const [openMode, override] = parseOverride(openModeIn);
                        const settings = {...plugin.settings, ...override};
                        const activeLeaf = this.getActiveViewOfType(View)?.leaf;

                        // Preview tab: intercept default (non-split, non-window) opens from file explorer
                        const shouldPreview = (
                            plugin.settings.previewTabs &&
                            plugin.nextOpenIsPreview &&
                            !openMode
                        );
                        plugin.nextOpenIsPreview = false; // consume the flag

                        let leaf: WorkspaceLeaf;
                        if (shouldPreview) {
                            // Find existing preview leaf globally
                            let existingPreview: WorkspaceLeaf | undefined;
                            if (plugin.currentPreviewLeafId) {
                                const found = this.getLeafById(plugin.currentPreviewLeafId);
                                if (found && isMainLeaf(found) && !isEmptyLeaf(found)) {
                                    existingPreview = found;
                                }
                            }
                            if (existingPreview) {
                                leaf = existingPreview;
                                // Activate the reused preview leaf so the user sees the new file.
                                // New leaves are activated inside createNewLeaf, but reused ones are not.
                                this.setActiveLeaf(existingPreview);
                            } else {
                                leaf = plugin.createNewLeaf(true, settings);
                            }
                        } else if (openMode == 'tab' || (!openMode && settings.openInNewTab)) {
                            // Tabs opened via normal click are always focused regardless of focusNewTab setting.
                            leaf = plugin.createNewLeaf(!openMode ? true : undefined, settings);
                        } else if (!openMode) {
                            leaf = plugin.getUnpinnedLeaf(true, settings);
                        } else {
                            leaf = (oldMethod as (...args: unknown[]) => WorkspaceLeaf).call(this, openMode, ...args);
                        }

                        // we set these to be used in openFile so we can tell when to deduplicate files.
                        leaf.openTabSettings = {
                            openMode, override,
                            openedFrom: activeLeaf?.id,
                            isPreview: shouldPreview,
                        }

                        return leaf;
                    }
                },

                /**
                 * getUnpinnedLeaf is deprecated in favor of getLeaf(false). However, it is used in a couple places in
                 * Obsidian and many plugins still use it directly. So we'll patch it as well to enforce new tab behavior.
                 *
                 * Note that as of 1.9.10, getUnpinnedLeaf takes an undocumented "focus" boolean. Obsidian uses this param
                 * when using ctrl and arrow keys in the file explorer to open files.
                 */
                getUnpinnedLeaf: (_oldMethod) => {
                    return function(this: Workspace, focus?: boolean) {
                        if (plugin.settings.openInNewTab) {
                            return this.getLeaf("tab");
                        } else {
                            return plugin.getUnpinnedLeaf(focus);
                        }
                    }
                },
            }));

            // Patch openFile to deduplicate tabs
            plugin.register(monkeyAround.around(WorkspaceLeaf.prototype, {
                openFile: (oldMethod) => {
                    return async function(this: WorkspaceLeaf, file: TFile, openState: Record<string, unknown>, ...args: unknown[]): Promise<void> {
                        let match: WorkspaceLeaf|undefined;

                        // these values are only valid immediately after creating a leaf. We clear them after openFile,
                        // and also clear them here if the leaf somehow gets populated without openFile.
                        // Read values BEFORE deleting — preview tab reuse sets openTabSettings on a non-empty leaf.
                        const tabSettings = this.openTabSettings;
                        if (!isEmptyLeaf(this)) delete this.openTabSettings;

                        const {openMode, override, openedFrom, isPreview} = tabSettings ?? {};
                        const settings = {...plugin.settings, ...override};

                        let matches = plugin.findMatchingLeaves(file);
                        if (!settings.deduplicateAcrossTabGroups) {
                            matches = matches.filter(l => l.parent == this.parent);
                        }
                        // When opening in preview mode, exclude existing preview tabs from dedup matches.
                        // We want to find permanent tabs only — preview tabs should be reused, not deduped to.
                        if (isPreview) {
                            matches = matches.filter(l => !plugin.isPreviewTab(l.id));
                        }

                        // if leaf is new and was opened via an explicit open in new window, split, or "allow duplicate",
                        // don't deduplicate. Note that opening in new window doesn't call getLeaf (it calls openPopoutLeaf
                        // directly) so we assume undefined openType is a new window. getLeaf("same") will update openType,
                        // so we shouldn't need to worry about if openType is undefined because the leaf was created before
                        // the plugin was loaded or such.
                        const isSpecialOpen = (
                            !isMainLeaf(this) ||
                            (isEmptyLeaf(this) && ![false, "tab"].includes(openMode ?? 'unknown'))
                        );
                        const eState = openState?.eState as Record<string, unknown> | undefined;
                        const isInternalLink = (
                            isEmptyLeaf(this) && openMode === false &&
                            !!eState?.subpath &&
                            matches.some(l => l.id == openedFrom)
                        );
                        const isMatch = matches.includes(this);

                        // if the link opened was an internal link, always deduplicate to undo open in new tab.
                        if (isInternalLink && !isSpecialOpen && !isMatch) {
                            match = matches.find(l => l.id == openedFrom)!;
                        } else if (settings.deduplicateTabs && !isSpecialOpen && matches.length > 0 && !isMatch) {
                            // choose matches first from last opened from, then matches in same group, then fist in list.
                            match = matches.find(l => l.id == openedFrom);
                            if (!match) matches.find(l => l.parent == this.parent);
                            if (!match) match = matches[0];
                        }

                        if (match) {
                            // If dedup routes to a preview tab (from a non-file-explorer source), promote it
                            if (plugin.isPreviewTab(match.id)) {
                                plugin.promotePreviewTab(match.id);
                            }
                            if (match.view.getViewType() == "kanban") {
                                // workaround for a bug in kanban. See
                                //     https://github.com/jesse-r-s-hines/obsidian-open-tab-settings/issues/25
                                //     https://github.com/mgmeyers/obsidian-kanban/issues/1102
                                plugin.app.workspace.setActiveLeaf(matches[0]);
                            } else {
                                const activeLeaf = plugin.app.workspace.getActiveViewOfType(View)?.leaf;
                                const shouldBeActive = !!openState?.active || activeLeaf == this;
                                await (oldMethod as (...args: unknown[]) => Promise<void>).call(matches[0], file, {
                                    ...openState,
                                    active: shouldBeActive,
                                }, ...args);
                                // Explicitly activate — openFile may be a no-op when the file is already
                                // showing, skipping the active flag processing.
                                if (shouldBeActive) {
                                    plugin.app.workspace.setActiveLeaf(matches[0]);
                                }
                            }
                        } else { // use default behavior
                            await (oldMethod as (...args: unknown[]) => Promise<void>).call(this, file, openState, ...args);
                        }

                        // Preview tab bookkeeping
                        if (isPreview && !match) {
                            plugin.markAsPreview(this);
                        }

                        // If the leaf is still empty, close it. This can happen if the file was de-duplicated while
                        // "openInNewTab" is enabled, or if you open a file "in default app" in a new tab.
                        if (isEmptyLeaf(this) && this.parent.children.length > 1) {
                            const tabGroup = this.parent;
                            const wasCurrentTab = tabGroup.children[tabGroup.currentTab] === this;
                            const lastActiveTab = tabGroup.children
                                .filter(l => l !== this)
                                .reduce((max, l) => l.activeTime > max.activeTime ? l : max);
                            this.detach();
                            if (wasCurrentTab) {
                                tabGroup.selectTabIndex(tabGroup.children.findIndex(c => c === lastActiveTab));
                            }
                        }

                        delete this.openTabSettings;
                    }
                },
            }));

            // Patch isModEvent to add override settings
            // We could have used isModEvent to implement openInNewTab instead of getLeaf, but there's quite a few places
            // that call getLeaf without isModEvent, such as the graph view.
            plugin.register(monkeyAround.around(Keymap, {
                isModEvent: (oldMethod) => {
                    return function(this: typeof Keymap, ...args: unknown[]): boolean | PaneType {
                        let result: boolean | PaneType | null = (oldMethod as (...args: unknown[]) => boolean | PaneType | null).call(this, ...args);
                        if (result == "tab") {
                            result = OVERRIDES[plugin.settings.modClickBehavior] as PaneType;
                        }
                        return result ?? false;
                    } as typeof Keymap.isModEvent;
                },
            }));
        };

        registerPatches(this);
    }

    onunload() {
        this.removePreviewTabHandlers();
        document.querySelectorAll('.is-preview-tab').forEach(el => {
            el.classList.remove('is-preview-tab');
        });
    }

    async loadSettings() {
        const dataFile = await this.loadData() ?? {};
        this.settings = Object.assign({}, DEFAULT_SETTINGS, dataFile);

        if (Object.keys(dataFile).length == 0) {
            // when using this plugin, focusNewTab should default to false. Set it if this is the first time we've
            // loaded the plugin.
            this.app.vault.setConfig('focusNewTab', false);
            await this.updateSettings({});
        }
    }

    async updateSettings(settings: Partial<OpenTabSettingsPluginSettings>) {
        const wasPreviewEnabled = this.settings.previewTabs;
        Object.assign(this.settings, settings);
        await this.saveData(this.settings);

        // Toggle preview tab handlers when the setting changes
        if (settings.previewTabs !== undefined && settings.previewTabs !== wasPreviewEnabled) {
            if (settings.previewTabs) {
                this.registerPreviewTabHandlers();
            } else {
                this.removePreviewTabHandlers();
                // Promote existing preview tab
                if (this.currentPreviewLeafId) {
                    this.promotePreviewTab(this.currentPreviewLeafId);
                }
            }
        }
    }

    private findMatchingLeaves(file: TFile) {
        const matches: WorkspaceLeaf[] = [];
        this.app.workspace.iterateAllLeaves(leaf => {
            // file is the same
            const isFileMatch = leaf.getViewState()?.state?.file == file.path;
            // we only want to switch to another leaf if its a basic file, not if its outgoing-links etc.
            const viewType = leaf.view.getViewType();
            const isTypeMatch = (
                this.app.viewRegistry.getTypeByExtension(file.extension) == viewType ||
                PLUGIN_VIEW_TYPES[file.extension]?.includes(viewType)
            );

            if (isMainLeaf(leaf) && isFileMatch && isTypeMatch) {
                matches.push(leaf);
            }
        });
        return matches;
    }

    /**
     * Gets all tab groups, sorted by active time.
     */
    private getAllTabGroups(root: WorkspaceItem): TabGroup[] {
        const tabGroups: Set<TabGroup> = new Set(); // sets are ordered
        this.app.workspace.iterateAllLeaves(leaf => {
            if (leaf.getRoot() == root) {
                tabGroups.add(leaf.parent);
            }
        });
        return [...tabGroups];
    }

    /**
     * Custom variant of the internal workspace.createLeafInTabGroup function that follows our new tab placement logic.
     * @param focus Whether to focus the new tab. If undefined focus based on focusNewTab config
     */
    private createNewLeaf(focus?: boolean, override: Partial<OpenTabSettingsPluginSettings> = {}) {
        const workspace = this.app.workspace;
        focus = focus ?? this.app.vault.getConfig('focusNewTab') as boolean;
        const settings = {...this.settings, ...override};

        const activeLeaf = workspace.getMostRecentLeaf();
        if (!activeLeaf) throw new Error("No tab group found.");
        const activeTabGroup = activeLeaf.parent;
        const activeIndex = activeTabGroup.children.indexOf(activeLeaf);

        // This is default Obsidian behavior, if active leaf is empty new tab replaces it instead of making a new one.
        if (isEmptyLeaf(activeLeaf)) {
            return activeLeaf;
        }

        let group: TabGroup|undefined;
        let index: number|undefined;

        if (settings.newTabTabGroupPlacement != "same" && !Platform.isPhone) {
            const tabGroups = this.getAllTabGroups(activeLeaf.getRoot());
            const otherTabGroup = tabGroups.filter(g => g !== activeTabGroup).at(-1);
            if (settings.newTabTabGroupPlacement == "opposite" && otherTabGroup) {
                group = otherTabGroup;
            } else if (settings.newTabTabGroupPlacement == "first" && tabGroups.at(0)) {
                group = tabGroups[0];
            } else if (settings.newTabTabGroupPlacement == "last" && tabGroups.at(-1)) {
                group = tabGroups.at(-1)!;
            }
        }
        if (!group) {
            group = activeTabGroup;
        }

        if (group == activeTabGroup) {
            if (settings.newTabPlacement == "after-pinned") {
                const lastPinnedIndex = group.children.findLastIndex(l => l.pinned);
                index = lastPinnedIndex >= 0 ? lastPinnedIndex + 1 : activeIndex + 1;
            } else if (settings.newTabPlacement == "beginning") {
                index = 0;
            } else if (settings.newTabPlacement == "end") {
                index = activeTabGroup.children.length;
            } else {
                index = activeIndex + 1;
            }
        } else {
            if (settings.newTabPlacement == "beginning") {
                index = 0
            } else {
                index = activeTabGroup.children.length;
            }
        }

        let newLeaf: WorkspaceLeaf;
        // we re-use empty tabs more aggressively than default Obsidian. If the tab at the new location is empty, re-use
        // it instead of creating a new one.
        const leafToDisplace = group.children[Math.min(index, group.children.length - 1)];
        if (isEmptyLeaf(leafToDisplace)) {
            newLeaf = leafToDisplace;
        } else {
            newLeaf = new (WorkspaceLeaf as unknown as new (app: App) => WorkspaceLeaf)(this.app);
            const currentTab = group.currentTab;
            // If new tab is inserted before the currently tab in a group, and we aren't setting the new tab active, we
            // need to update the selected tab so that group.currentTab index still points to the original active tab
            group.insertChild(index, newLeaf);
            if (index <= currentTab && (group != activeTabGroup || !focus)) {
                group.selectTabIndex(currentTab + 1);
            }
        }

        if (focus) {
            workspace.setActiveLeaf(newLeaf);
        }

        return newLeaf;
    }

    /**
     * Custom implementation of getUnpinnedLeaf that implements our new tab placement behavior when making new tabs,
     * e.g. when the active tab is pinned.
     */
    private getUnpinnedLeaf(focus = true, override: Partial<OpenTabSettingsPluginSettings> = {}) {
        const workspace = this.app.workspace;
        const settings = {...this.settings, ...override};

        const activeLeaf = workspace.getActiveViewOfType(View)?.leaf;
        if (activeLeaf?.canNavigate()) {
            return activeLeaf;
        }

        const container = activeLeaf?.getContainer() ?? workspace.rootSplit;

        let leaf: WorkspaceLeaf|null = null;
        workspace.iterateLeaves(container, (l) => {
          if (l.canNavigate()) {
            const group = l.parent;
            if (
                group &&
                (group.children[group.currentTab] === l || (group instanceof WorkspaceTabs && group.isStacked)) &&
                (!leaf || leaf.activeTime < l.activeTime)
            ) {
              leaf = l;
            }
          }
        });

        if (!leaf) {
            leaf = this.createNewLeaf(focus, settings);
        } else if (focus) {
            workspace.setActiveLeaf(leaf);
        }

        return leaf;
    }

    // --- Preview Tab Methods ---

    private registerPreviewTabHandlers() {
        if (this.previewHandlersRegistered || !this.settings.previewTabs) return;
        this.previewHandlersRegistered = true;

        // Click handler: detect file explorer clicks and set preview flag
        this.boundClickHandler = (evt: MouseEvent) => {
            if (!this.settings.previewTabs) return;
            const target = (evt.target as HTMLElement).closest('.nav-file-title');
            if (!target || !target.getAttribute('data-path')) return;
            // Ignore modifier clicks — those go through existing mod-click behavior
            if (evt.button !== 0 || evt.ctrlKey || evt.metaKey || evt.shiftKey || evt.altKey) return;

            // Don't preview if file is already open in a permanent tab — let normal dedup handle it.
            // Reusing the preview leaf for a different file that's already open elsewhere causes
            // openFile to be a no-op on the matched leaf (same file already showing).
            const filePath = target.getAttribute('data-path')!;
            const file = this.app.vault.getAbstractFileByPath(filePath);
            if (file instanceof TFile) {
                const permanentMatches = this.findMatchingLeaves(file).filter(l => !this.isPreviewTab(l.id));
                if (permanentMatches.length > 0) {
                    return;
                }
            }

            this.nextOpenIsPreview = true;
        };
        document.addEventListener('click', this.boundClickHandler, true);

        // Double-click handler: promote the preview tab to permanent
        this.boundDblClickHandler = (evt: MouseEvent) => {
            if (!this.settings.previewTabs) return;
            const target = (evt.target as HTMLElement).closest('.nav-file-title');
            if (!target) return;
            const filePath = target.getAttribute('data-path');
            if (!filePath) return;

            this.promotePreviewForFile(filePath);
        };
        document.addEventListener('dblclick', this.boundDblClickHandler, true);

        // Auto-promote on edit
        this.registerEvent(
            this.app.workspace.on('editor-change', () => {
                if (!this.settings.previewTabsAutoPromote || !this.currentPreviewLeafId) return;
                const activeLeaf = this.app.workspace.getActiveViewOfType(View)?.leaf;
                if (activeLeaf && activeLeaf.id === this.currentPreviewLeafId) {
                    this.promotePreviewTab(activeLeaf.id);
                }
            })
        );

        // Clean up preview tracking on layout changes (tab closed, pinned, etc.)
        this.registerEvent(
            this.app.workspace.on('layout-change', () => {
                if (this.currentPreviewLeafId) {
                    const leaf = this.app.workspace.getLeafById(this.currentPreviewLeafId);
                    if (!leaf) {
                        this.currentPreviewLeafId = null;
                    } else if (leaf.pinned) {
                        this.promotePreviewTab(this.currentPreviewLeafId);
                    }
                }
            })
        );
    }

    private removePreviewTabHandlers() {
        if (this.boundClickHandler) {
            document.removeEventListener('click', this.boundClickHandler, true);
            this.boundClickHandler = null;
        }
        if (this.boundDblClickHandler) {
            document.removeEventListener('dblclick', this.boundDblClickHandler, true);
            this.boundDblClickHandler = null;
        }
        this.previewHandlersRegistered = false;
    }

    private isPreviewTab(leafId: string): boolean {
        return this.currentPreviewLeafId === leafId;
    }

    private markAsPreview(leaf: WorkspaceLeaf) {
        // Clear previous preview style
        if (this.currentPreviewLeafId && this.currentPreviewLeafId !== leaf.id) {
            this.updatePreviewStyle(this.currentPreviewLeafId, false);
        }
        this.currentPreviewLeafId = leaf.id;
        // Defer style update — tab header DOM may not exist yet
        requestAnimationFrame(() => this.updatePreviewStyle(leaf.id, true));
    }

    private promotePreviewTab(leafId: string) {
        if (this.currentPreviewLeafId === leafId) {
            this.currentPreviewLeafId = null;
        }
        this.updatePreviewStyle(leafId, false);
    }

    private promotePreviewForFile(filePath: string) {
        if (this.currentPreviewLeafId) {
            const leaf = this.app.workspace.getLeafById(this.currentPreviewLeafId);
            if (leaf && leaf.getViewState()?.state?.file === filePath) {
                this.promotePreviewTab(this.currentPreviewLeafId);
                return;
            }
        }
    }

    private updatePreviewStyle(leafId: string, isPreview: boolean) {
        const leaf = this.app.workspace.getLeafById(leafId);
        const tabHeader = (leaf as unknown as { readonly tabHeaderEl?: HTMLElement })?.tabHeaderEl;
        if (tabHeader) {
            tabHeader.classList.toggle('is-preview-tab', isPreview);
        }
    }
}
