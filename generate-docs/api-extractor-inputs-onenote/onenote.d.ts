////////////////////////////////////////////////////////////////
////////////////////// Begin OneNote APIs //////////////////////
////////////////////////////////////////////////////////////////


export declare namespace OneNote {
    /**
     *
     * Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Application extends OfficeExtension.ClientObject {
        private m_notebooks;
        /**
         *
         * Gets the collection of notebooks that are open in the OneNote application instance. In OneNote Online, only one notebook at a time is open in the application instance. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        notebooks: OneNote.NotebookCollection;
        /**
         *
         * Gets the active notebook if one exists. If no notebook is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveNotebook(): OneNote.Notebook;
        /**
         *
         * Gets the active notebook if one exists. If no notebook is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveNotebookOrNull(): OneNote.Notebook;
        /**
         *
         * Gets the active outline if one exists, If no outline is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveOutline(): OneNote.Outline;
        /**
         *
         * Gets the active outline if one exists, otherwise returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveOutlineOrNull(): OneNote.Outline;
        /**
         *
         * Gets the active page if one exists. If no page is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActivePage(): OneNote.Page;
        /**
         *
         * Gets the active page if one exists. If no page is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActivePageOrNull(): OneNote.Page;
        /**
         *
         * Gets the active section if one exists. If no section is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveSection(): OneNote.Section;
        /**
         *
         * Gets the active section if one exists. If no section is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveSectionOrNull(): OneNote.Section;
        /**
         *
         * Opens the specified page in the application instance.
         *
         * @param page - The page to open.
         *
         * [Api set: OneNoteApi 1.1]
         */
        navigateToPage(page: OneNote.Page): void;
        /**
         *
         * Gets the specified page, and opens it in the application instance.
         *
         * @param url - The client url of the page to open.
         *
         * [Api set: OneNoteApi 1.1]
         */
        navigateToPageWithClientUrl(url: string): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Application;
    }
    /**
     *
     * Represents ink analysis data for a given set of ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysis extends OfficeExtension.ClientObject {
        private m_id;
        private m_page;
        private m_paragraphs;
        private m__ReferenceId;
        /**
         *
         * Gets the parent page object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        page: OneNote.Page;
        /**
         *
         * Gets the ink analysis paragraphs in this page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraphs: OneNote.InkAnalysisParagraphCollection;
        /**
         *
         * Gets the ID of the InkAnalysis object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysis;
    }
    /**
     *
     * Represents ink analysis data for an identified paragraph formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisParagraph extends OfficeExtension.ClientObject {
        private m_id;
        private m_inkAnalysis;
        private m_lines;
        private m__ReferenceId;
        /**
         *
         * Reference to the parent InkAnalysisPage. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        inkAnalysis: OneNote.InkAnalysis;
        /**
         *
         * Gets the ink analysis lines in this ink analysis paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        lines: OneNote.InkAnalysisLineCollection;
        /**
         *
         * Gets the ID of the InkAnalysisParagraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisParagraph;
    }
    /**
     *
     * Represents a collection of InkAnalysisParagraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisParagraphCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.InkAnalysisParagraph>;
        /**
         *
         * Returns the number of InkAnalysisParagraphs in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a InkAnalysisParagraph object by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the InkAnalysisParagraph object, or the index location of the InkAnalysisParagraph object in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.InkAnalysisParagraph;
        /**
         *
         * Gets a InkAnalysisParagraph on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.InkAnalysisParagraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisParagraphCollection;
    }
    /**
     *
     * Represents ink analysis data for an identified text line formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisLine extends OfficeExtension.ClientObject {
        private m_id;
        private m_paragraph;
        private m_words;
        private m__ReferenceId;
        /**
         *
         * Reference to the parent InkAnalysisParagraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraph: OneNote.InkAnalysisParagraph;
        /**
         *
         * Gets the ink analysis words in this ink analysis line. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        words: OneNote.InkAnalysisWordCollection;
        /**
         *
         * Gets the ID of the InkAnalysisLine object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisLine;
    }
    /**
     *
     * Represents a collection of InkAnalysisLine objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisLineCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.InkAnalysisLine>;
        /**
         *
         * Returns the number of InkAnalysisLines in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a InkAnalysisLine object by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the InkAnalysisLine object, or the index location of the InkAnalysisLine object in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.InkAnalysisLine;
        /**
         *
         * Gets a InkAnalysisLine on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.InkAnalysisLine;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisLineCollection;
    }
    /**
     *
     * Represents ink analysis data for an identified word formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisWord extends OfficeExtension.ClientObject {
        private m_id;
        private m_languageId;
        private m_line;
        private m_strokePointers;
        private m_wordAlternates;
        private m__ReferenceId;
        /**
         *
         * Reference to the parent InkAnalysisLine. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        line: OneNote.InkAnalysisLine;
        /**
         *
         * Gets the ID of the InkAnalysisWord object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * The id of the recognized language in this inkAnalysisWord. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        languageId: string;
        /**
         *
         * Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        strokePointers: Array<OneNote.InkStrokePointer>;
        /**
         *
         * The words that were recognized in this ink word, in order of likelihood. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        wordAlternates: Array<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisWord;
    }
    /**
     *
     * Represents a collection of InkAnalysisWord objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisWordCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.InkAnalysisWord>;
        /**
         *
         * Returns the number of InkAnalysisWords in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a InkAnalysisWord object by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the InkAnalysisWord object, or the index location of the InkAnalysisWord object in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.InkAnalysisWord;
        /**
         *
         * Gets a InkAnalysisWord on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.InkAnalysisWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkAnalysisWordCollection;
    }
    /**
     *
     * Represents a group of ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class FloatingInk extends OfficeExtension.ClientObject {
        private m_id;
        private m_inkStrokes;
        private m_pageContent;
        private m__ReferenceId;
        /**
         *
         * Gets the strokes of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        inkStrokes: OneNote.InkStrokeCollection;
        /**
         *
         * Gets the PageContent parent of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        pageContent: OneNote.PageContent;
        /**
         *
         * Gets the ID of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.FloatingInk;
    }
    /**
     *
     * Represents a single stroke of ink.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkStroke extends OfficeExtension.ClientObject {
        private m_floatingInk;
        private m_id;
        private m__ReferenceId;
        /**
         *
         * Gets the ID of the InkStroke object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        floatingInk: OneNote.FloatingInk;
        /**
         *
         * Gets the ID of the InkStroke object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkStroke;
    }
    /**
     *
     * Represents a collection of InkStroke objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkStrokeCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.InkStroke>;
        /**
         *
         * Returns the number of InkStrokes in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a InkStroke object by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the InkStroke object, or the index location of the InkStroke object in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.InkStroke;
        /**
         *
         * Gets a InkStroke on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.InkStroke;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkStrokeCollection;
    }
    /**
     *
     * A container for the ink in a word in a paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkWord extends OfficeExtension.ClientObject {
        private m_id;
        private m_languageId;
        private m_paragraph;
        private m_wordAlternates;
        private m__ReferenceId;
        /**
         *
         * The parent paragraph containing the ink word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraph: OneNote.Paragraph;
        /**
         *
         * Gets the ID of the InkWord object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * The id of the recognized language in this ink word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        languageId: string;
        /**
         *
         * The words that were recognized in this ink word, in order of likelihood. Read-only.
         *
         * [Api set: OneNoteApi]
         */
        wordAlternates: Array<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkWord;
    }
    /**
     *
     * Represents a collection of InkWord objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkWordCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.InkWord>;
        /**
         *
         * Returns the number of InkWords in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a InkWord object by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the InkWord object, or the index location of the InkWord object in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.InkWord;
        /**
         *
         * Gets a InkWord on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.InkWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.InkWordCollection;
    }
    /**
     *
     * Represents a OneNote notebook. Notebooks contain section groups and sections.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Notebook extends OfficeExtension.ClientObject {
        private m_clientUrl;
        private m_id;
        private m_name;
        private m_sectionGroups;
        private m_sections;
        private m__ReferenceId;
        /**
         *
         * The section groups in the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        sectionGroups: OneNote.SectionGroupCollection;
        /**
         *
         * The the sections of the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        sections: OneNote.SectionCollection;
        /**
         *
         * The client url of the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        clientUrl: string;
        /**
         *
         * Gets the ID of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the name of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        name: string;
        /**
         *
         * Adds a new section to the end of the notebook.
         *
         * @param name - The name of the new section.
         *
         * [Api set: OneNoteApi 1.1]
         */
        addSection(name: string): OneNote.Section;
        /**
         *
         * Adds a new section group to the end of the notebook.
         *
         * @param name - The name of the new section.
         *
         * [Api set: OneNoteApi 1.1]
         */
        addSectionGroup(name: string): OneNote.SectionGroup;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Notebook;
    }
    /**
     *
     * Represents a collection of notebooks.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class NotebookCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.Notebook>;
        /**
         *
         * Returns the number of notebooks in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets the collection of notebooks with the specified name that are open in the application instance.
         *
         * @param name - The name of the notebook.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getByName(name: string): OneNote.NotebookCollection;
        /**
         *
         * Gets a notebook by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the notebook, or the index location of the notebook in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.Notebook;
        /**
         *
         * Gets a notebook on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.Notebook;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.NotebookCollection;
    }
    /**
     *
     * Represents a OneNote section group. Section groups can contain sections and other section groups.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class SectionGroup extends OfficeExtension.ClientObject {
        private m_clientUrl;
        private m_id;
        private m_name;
        private m_notebook;
        private m_parentSectionGroup;
        private m_parentSectionGroupOrNull;
        private m_sectionGroups;
        private m_sections;
        private m__ReferenceId;
        /**
         *
         * Gets the notebook that contains the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        notebook: OneNote.Notebook;
        /**
         *
         * Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentSectionGroup: OneNote.SectionGroup;
        /**
         *
         * Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentSectionGroupOrNull: OneNote.SectionGroup;
        /**
         *
         * The collection of section groups in the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        sectionGroups: OneNote.SectionGroupCollection;
        /**
         *
         * The collection of sections in the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        sections: OneNote.SectionCollection;
        /**
         *
         * The client url of the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        clientUrl: string;
        /**
         *
         * Gets the ID of the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the name of the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        name: string;
        /**
         *
         * Adds a new section to the end of the section group.
         *
         * @param title - The name of the new section.
         *
         * [Api set: OneNoteApi 1.1]
         */
        addSection(title: string): OneNote.Section;
        /**
         *
         * Adds a new section group to the end of this sectionGroup.
         *
         * @param name - The name of the new section.
         *
         * [Api set: OneNoteApi 1.1]
         */
        addSectionGroup(name: string): OneNote.SectionGroup;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.SectionGroup;
    }
    /**
     *
     * Represents a collection of section groups.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class SectionGroupCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.SectionGroup>;
        /**
         *
         * Returns the number of section groups in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets the collection of section groups with the specified name.
         *
         * @param name - The name of the section group.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getByName(name: string): OneNote.SectionGroupCollection;
        /**
         *
         * Gets a section group by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the section group, or the index location of the section group in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.SectionGroup;
        /**
         *
         * Gets a section group on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.SectionGroup;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.SectionGroupCollection;
    }
    /**
     *
     * Represents a OneNote section. Sections can contain pages.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Section extends OfficeExtension.ClientObject {
        private m_clientUrl;
        private m_id;
        private m_name;
        private m_notebook;
        private m_pages;
        private m_parentSectionGroup;
        private m_parentSectionGroupOrNull;
        private m__ReferenceId;
        /**
         *
         * Gets the notebook that contains the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        notebook: OneNote.Notebook;
        /**
         *
         * The collection of pages in the section. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        pages: OneNote.PageCollection;
        /**
         *
         * Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentSectionGroup: OneNote.SectionGroup;
        /**
         *
         * Gets the section group that contains the section. Returns null if the section is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentSectionGroupOrNull: OneNote.SectionGroup;
        /**
         *
         * The client url of the section. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        clientUrl: string;
        /**
         *
         * Gets the ID of the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the name of the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        name: string;
        /**
         *
         * Adds a new page to the end of the section.
         *
         * @param title - The title of the new page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        addPage(title: string): OneNote.Page;
        /**
         *
         * Copies this section to specified notebook.
         *
         * @param destinationNotebook - The notebook to copy this section to.
         *
         * [Api set: OneNoteApi 1.1]
         */
        copyToNotebook(destinationNotebook: OneNote.Notebook): OneNote.Section;
        /**
         *
         * Copies this section to specified section group.
         *
         * @param destinationSectionGroup - The section group to copy this section to.
         *
         * [Api set: OneNoteApi 1.1]
         */
        copyToSectionGroup(destinationSectionGroup: OneNote.SectionGroup): OneNote.Section;
        /**
         *
         * Inserts a new section before or after the current section.
         *
         * @param location - The location of the new section relative to the current section.
         * @param title - The name of the new section.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertSectionAsSibling(location: string, title: string): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Section;
    }
    /**
     *
     * Represents a collection of sections.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class SectionCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.Section>;
        /**
         *
         * Returns the number of sections in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets the collection of sections with the specified name.
         *
         * @param name - The name of the section.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getByName(name: string): OneNote.SectionCollection;
        /**
         *
         * Gets a section by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the section, or the index location of the section in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.Section;
        /**
         *
         * Gets a section on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.SectionCollection;
    }
    /**
     *
     * Represents a OneNote page.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Page extends OfficeExtension.ClientObject {
        private m_clientUrl;
        private m_contents;
        private m_id;
        private m_inkAnalysisOrNull;
        private m_pageLevel;
        private m_parentSection;
        private m_title;
        private m_webUrl;
        private m__ReferenceId;
        /**
         *
         * The collection of PageContent objects on the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        contents: OneNote.PageContentCollection;
        /**
         *
         * Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        inkAnalysisOrNull: OneNote.InkAnalysis;
        /**
         *
         * Gets the section that contains the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentSection: OneNote.Section;
        /**
         *
         * The client url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        clientUrl: string;
        /**
         *
         * Gets the ID of the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets or sets the indentation level of the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        pageLevel: number;
        /**
         *
         * Gets or sets the title of the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        title: string;
        /**
         *
         * The web url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        webUrl: string;
        /**
         *
         * Adds an Outline to the page at the specified position.
         *
         * @param left - The left position of the top, left corner of the Outline.
         * @param top - The top position of the top, left corner of the Outline.
         * @param html - An HTML string that describes the visual presentation of the Outline. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         *
         * [Api set: OneNoteApi 1.1]
         */
        addOutline(left: number, top: number, html: string): OneNote.Outline;
        /**
         *
         * Copies this page to specified section.
         *
         * @param destinationSection - The section to copy this page to.
         *
         * [Api set: OneNoteApi 1.1]
         */
        copyToSection(destinationSection: OneNote.Section): OneNote.Page;
        /**
         *
         * Inserts a new page before or after the current page.
         *
         * @param location - The location of the new page relative to the current page.
         * @param title - The title of the new page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertPageAsSibling(location: string, title: string): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Page;
    }
    /**
     *
     * Represents a collection of pages.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class PageCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.Page>;
        /**
         *
         * Returns the number of pages in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets the collection of pages with the specified title.
         *
         * @param title - The title of the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getByTitle(title: string): OneNote.PageCollection;
        /**
         *
         * Gets a page by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the page, or the index location of the page in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.Page;
        /**
         *
         * Gets a page on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.PageCollection;
    }
    /**
     *
     * Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class PageContent extends OfficeExtension.ClientObject {
        private m_id;
        private m_image;
        private m_ink;
        private m_left;
        private m_outline;
        private m_parentPage;
        private m_top;
        private m_type;
        private m__ReferenceId;
        /**
         *
         * Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        image: OneNote.Image;
        /**
         *
         * Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
         *
         * [Api set: OneNoteApi 1.1]
         */
        ink: OneNote.FloatingInk;
        /**
         *
         * Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
         *
         * [Api set: OneNoteApi 1.1]
         */
        outline: OneNote.Outline;
        /**
         *
         * Gets the page that contains the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentPage: OneNote.Page;
        /**
         *
         * Gets the ID of the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets or sets the left (X-axis) position of the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        left: number;
        /**
         *
         * Gets or sets the top (Y-axis) position of the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        top: number;
        /**
         *
         * Gets the type of the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        type: string;
        /**
         *
         * Deletes the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.PageContent;
    }
    /**
     *
     * Represents the contents of a page, as a collection of PageContent objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class PageContentCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.PageContent>;
        /**
         *
         * Returns the number of page contents in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a PageContent object by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the PageContent object, or the index location of the PageContent object in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.PageContent;
        /**
         *
         * Gets a page content on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.PageContent;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.PageContentCollection;
    }
    /**
     *
     * Represents a container for Paragraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Outline extends OfficeExtension.ClientObject {
        private m_id;
        private m_pageContent;
        private m_paragraphs;
        private m__ReferenceId;
        /**
         *
         * Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        pageContent: OneNote.PageContent;
        /**
         *
         * Gets the collection of Paragraph objects in the Outline. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the ID of the Outline object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Adds the specified HTML to the bottom of the Outline.
         *
         * @param html - The HTML string to append. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendHtml(html: string): void;
        /**
         *
         * Adds the specified image to the bottom of the Outline.
         *
         * @param base64EncodedImage - HTML string to append.
         * @param width - Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height - Optional. Height in the unit of Points. The default value is null and image height will be respected.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendImage(base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Adds the specified text to the bottom of the Outline.
         *
         * @param paragraphText - HTML string to append.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendRichText(paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns to the bottom of the outline.
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendTable(rowCount: number, columnCount: number, values?: Array<Array<string>>): OneNote.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Outline;
    }
    /**
     *
     * A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Paragraph extends OfficeExtension.ClientObject {
        private m_id;
        private m_image;
        private m_inkWords;
        private m_outline;
        private m_paragraphs;
        private m_parentParagraph;
        private m_parentParagraphOrNull;
        private m_parentTableCell;
        private m_parentTableCellOrNull;
        private m_richText;
        private m_table;
        private m_type;
        private m__ReferenceId;
        /**
         *
         * Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        image: OneNote.Image;
        /**
         *
         * Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        inkWords: OneNote.InkWordCollection;
        /**
         *
         * Gets the Outline object that contains the Paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        outline: OneNote.Outline;
        /**
         *
         * The collection of paragraphs under this paragraph. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the parent paragraph object. Throws if a parent paragraph does not exist. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentParagraph: OneNote.Paragraph;
        /**
         *
         * Gets the parent paragraph object. Returns null if a parent paragraph does not exist. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentParagraphOrNull: OneNote.Paragraph;
        /**
         *
         * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentTableCell: OneNote.TableCell;
        /**
         *
         * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentTableCellOrNull: OneNote.TableCell;
        /**
         *
         * Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText. Read-only
         *
         * [Api set: OneNoteApi 1.1]
         */
        richText: OneNote.RichText;
        /**
         *
         * Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        table: OneNote.Table;
        /**
         *
         * Gets the ID of the Paragraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the type of the Paragraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        type: string;
        /**
         *
         * Deletes the paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        delete(): void;
        /**
         *
         * Inserts the specified HTML content
         *
         * @param insertLocation - The location of new contents relative to the current Paragraph.
         * @param html - An HTML string that describes the visual presentation of the content. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertHtmlAsSibling(insertLocation: string, html: string): void;
        /**
         *
         * Inserts the image at the specified insert location..
         *
         * @param insertLocation - The location of the table relative to the current Paragraph.
         * @param base64EncodedImage - HTML string to append.
         * @param width - Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height - Optional. Height in the unit of Points. The default value is null and image height will be respected.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Inserts the paragraph text at the specifiec insert location.
         *
         * @param insertLocation - The location of the table relative to the current Paragraph.
         * @param paragraphText - HTML string to append.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertRichTextAsSibling(insertLocation: string, paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns before or after the current paragraph.
         *
         * @param insertLocation - The location of the table relative to the current Paragraph.
         * @param rowCount - The number of rows in the table.
         * @param columnCount - The number of columns in the table.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values?: Array<Array<string>>): OneNote.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Paragraph;
    }
    /**
     *
     * Represents a collection of Paragraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class ParagraphCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.Paragraph>;
        /**
         *
         * Returns the number of paragraphs in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a Paragraph object by ID or by its index in the collection. Read-only.
         *
         * @param index - The ID of the Paragraph object, or the index location of the Paragraph object in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.Paragraph;
        /**
         *
         * Gets a paragraph on its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.ParagraphCollection;
    }
    /**
     *
     * Represents a RichText object in a Paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class RichText extends OfficeExtension.ClientObject {
        private m_id;
        private m_paragraph;
        private m_text;
        private m__ReferenceId;
        /**
         *
         * Gets the Paragraph object that contains the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraph: OneNote.Paragraph;
        /**
         *
         * Gets the ID of the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the text content of the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        text: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.RichText;
    }
    /**
     *
     * Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Image extends OfficeExtension.ClientObject {
        private m_description;
        private m_height;
        private m_hyperlink;
        private m_id;
        private m_ocrData;
        private m_pageContent;
        private m_paragraph;
        private m_width;
        private m__ReferenceId;
        /**
         *
         * Gets the PageContent object that contains the Image. Throws if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        pageContent: OneNote.PageContent;
        /**
         *
         * Gets the Paragraph object that contains the Image. Throws if the Image is not a direct child of a Paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraph: OneNote.Paragraph;
        /**
         *
         * Gets or sets the description of the Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        description: string;
        /**
         *
         * Gets or sets the height of the Image layout.
         *
         * [Api set: OneNoteApi 1.1]
         */
        height: number;
        /**
         *
         * Gets or sets the hyperlink of the Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        hyperlink: string;
        /**
         *
         * Gets the ID of the Image object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
         *
         * [Api set: OneNoteApi 1.1]
         */
        ocrData: OneNote.ImageOcrData;
        /**
         *
         * Gets or sets the width of the Image layout.
         *
         * [Api set: OneNoteApi 1.1]
         */
        width: number;
        /**
         *
         * Gets the base64-encoded binary representation of the Image.
            Example: data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIA...
         *
         * [Api set: OneNoteApi 1.1]
         */
        getBase64Image(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Image;
    }
    /**
     *
     * Represents a table in a OneNote page.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Table extends OfficeExtension.ClientObject {
        private m_borderVisible;
        private m_columnCount;
        private m_id;
        private m_paragraph;
        private m_rowCount;
        private m_rows;
        private m__ReferenceId;
        /**
         *
         * Gets the Paragraph object that contains the Table object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraph: OneNote.Paragraph;
        /**
         *
         * Gets all of the table rows. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        rows: OneNote.TableRowCollection;
        /**
         *
         * Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
         *
         * [Api set: OneNoteApi 1.1]
         */
        borderVisible: boolean;
        /**
         *
         * Gets the number of columns in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        columnCount: number;
        /**
         *
         * Gets the ID of the table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the number of rows in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        rowCount: number;
        /**
         *
         * Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.
         *
         * @param values - Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendColumn(values?: Array<string>): void;
        /**
         *
         * Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.
         *
         * @param values - Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendRow(values?: Array<string>): OneNote.TableRow;
        /**
         *
         * Clears the contents of the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         *
         * Gets the table cell at a specified row and column.
         *
         * @param rowIndex - The index of the row.
         * @param cellIndex - The index of the cell in the row.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getCell(rowIndex: number, cellIndex: number): OneNote.TableCell;
        /**
         *
         * Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.
         *
         * @param index - Index where the column will be inserted in the table.
         * @param values - Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertColumn(index: number, values?: Array<string>): void;
        /**
         *
         * Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.
         *
         * @param index - Index where the row will be inserted in the table.
         * @param values - Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertRow(index: number, values?: Array<string>): OneNote.TableRow;
        setShadingColor(colorCode: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.Table;
    }
    /**
     *
     * Represents a row in a table.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableRow extends OfficeExtension.ClientObject {
        private m_cellCount;
        private m_cells;
        private m_id;
        private m_parentTable;
        private m_rowIndex;
        private m__ReferenceId;
        /**
         *
         * Gets the cells in the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        cells: OneNote.TableCellCollection;
        /**
         *
         * Gets the parent table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentTable: OneNote.Table;
        /**
         *
         * Gets the number of cells in the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        cellCount: number;
        /**
         *
         * Gets the ID of the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the index of the row in its parent table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        rowIndex: number;
        /**
         *
         * Clears the contents of the row.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         *
         * Inserts a row before or after the current row.
         *
         * @param insertLocation - Where the new rows should be inserted relative to the current row.
         * @param values - Strings to insert in the new row, specified as an array. Must not have more cells than in the current row. Optional.
         *
         * [Api set: OneNoteApi 1.1]
         */
        insertRowAsSibling(insertLocation: string, values?: Array<string>): OneNote.TableRow;
        setShadingColor(colorCode: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableRow;
    }
    /**
     *
     * Contains a collection of TableRow objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableRowCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.TableRow>;
        /**
         *
         * Returns the number of table rows in this collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a table row object by ID or by its index in the collection. Read-only.
         *
         * @param index - A number that identifies the index location of a table row object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.TableRow;
        /**
         *
         * Gets a table row at its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableRowCollection;
    }
    /**
     *
     * Represents a cell in a OneNote table.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableCell extends OfficeExtension.ClientObject {
        private m_cellIndex;
        private m_id;
        private m_paragraphs;
        private m_parentRow;
        private m_rowIndex;
        private m_shadingColor;
        private m__ReferenceId;
        /**
         *
         * Gets the collection of Paragraph objects in the TableCell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the parent row of the cell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        parentRow: OneNote.TableRow;
        /**
         *
         * Gets the index of the cell in its row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        cellIndex: number;
        /**
         *
         * Gets the ID of the cell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        id: string;
        /**
         *
         * Gets the index of the cell's row in the table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        rowIndex: number;
        /**
         *
         * Gets and sets the shading color of the cell
         *
         * [Api set: OneNoteApi 1.1]
         */
        shadingColor: string;
        /**
         *
         * Adds the specified HTML to the bottom of the TableCell.
         *
         * @param html - The HTML string to append. See [supported HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html) for the OneNote add-ins JavaScript API.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendHtml(html: string): void;
        /**
         *
         * Adds the specified image to table cell.
         *
         * @param base64EncodedImage - HTML string to append.
         * @param width - Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height - Optional. Height in the unit of Points. The default value is null and image height will be respected.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendImage(base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Adds the specified text to table cell.
         *
         * @param paragraphText - HTML string to append.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendRichText(paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns to table cell.
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         *
         * [Api set: OneNoteApi 1.1]
         */
        appendTable(rowCount: number, columnCount: number, values?: Array<Array<string>>): OneNote.Table;
        /**
         *
         * Clears the contents of the cell.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableCell;
    }
    /**
     *
     * Contains a collection of TableCell objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableCellCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<OneNote.TableCell>;
        /**
         *
         * Returns the number of tablecells in this collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a table cell object by ID or by its index in the collection. Read-only.
         *
         * @param index - A number that identifies the index location of a table cell object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItem(index: number | string): OneNote.TableCell;
        /**
         *
         * Gets a tablecell at its position in the collection.
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getItemAt(index: number): OneNote.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OneNote.TableCellCollection;
    }
    /**
     *
     * Represents data obtained by OCR (optical character recognition) of an image
     *
     * [Api set: OneNoteApi 1.1]
     */
    export interface ImageOcrData {
        /**
         *
         * Represents the OCR language, with values such as EN-US
         *
         * [Api set: OneNoteApi 1.1]
         */
        ocrLanguageId: string;
        /**
         *
         * Represents the text obtained by OCR of the image
         *
         * [Api set: OneNoteApi 1.1]
         */
        ocrText: string;
    }
    /**
     *
     * Weak reference to an ink stroke object and its content parent
     *
     * [Api set: OneNoteApi 1.1]
     */
    export interface InkStrokePointer {
        /**
         *
         * Represents the id of the page content object corresponding to this stroke
         *
         * [Api set: OneNoteApi 1.1]
         */
        contentId: string;
        /**
         *
         * Represents the id of the ink stroke
         *
         * [Api set: OneNoteApi 1.1]
         */
        inkStrokeId: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    export module InsertLocation {
        var before: string;
        var after: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    export module Alignment {
        var left: string;
        var centered: string;
        var right: string;
        var justified: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    export module Selected {
        var notSelected: string;
        var partialSelected: string;
        var selected: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    export module PageContentType {
        var outline: string;
        var image: string;
        var ink: string;
        var other: string;
    }
    /**
     * [Api set: OneNoteApi]
     */
    export module ParagraphType {
        var richText: string;
        var image: string;
        var table: string;
        var ink: string;
        var other: string;
    }
    export module ErrorCodes {
        var generalException: string;
    }
}
export declare namespace OneNote {
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        private m_onenote;
        constructor(url?: string);
        application: Application;
    }
    /**
 * Executes a batch script that performs actions on the OneNote object model. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
 * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the OneNote application. Since the Office add-in and the WoOneNote application run in two different processes, the request context is required to get access to the OneNote object model from the add-in.
 */
    export function run<T>(batch: (context: OneNote.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}


////////////////////////////////////////////////////////////////
/////////////////////// End OneNote APIs ///////////////////////
////////////////////////////////////////////////////////////////