# OneNote JavaScript API overview

Applies to: OneNote Online

The following links show the high level OneNote objects available in the API. Each object page link contains a description of the properties, relationships, and methods available on the object. Explore these links to learn more. 
	
- [Application](../../docs-ref-autogen/onenote/onenote.application.yml): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.

- [Notebook](../../docs-ref-autogen/onenote/onenote.notebook.yml): A notebook. Notebooks contain section groups and sections.
    - [NotebookCollection](../../docs-ref-autogen/onenote/onenote.notebookcollection.yml): A collection of notebooks.

- [SectionGroup](../../docs-ref-autogen/onenote/onenote.sectiongroup.yml): A section group. Section groups contain section groups and sections.
    - [SectionGroupCollection](../../docs-ref-autogen/onenote/onenote.sectiongroupcollection.yml): A collection of section groups.

- [Section](../../docs-ref-autogen/onenote/onenote.section.yml): A section. Sections contain pages.
    - [SectionCollection](../../docs-ref-autogen/onenote/onenote.sectioncollection): A collection of sections.

- [Page](../../docs-ref-autogen/onenote/onenote.page.yml): A page. Pages contain PageContent objects.
    - [PageCollection](../../docs-ref-autogen/onenote/onenote.pagecollection.yml): A collection of pages.

- [PageContent](../../docs-ref-autogen/onenote/onenote.pagecontent.yml): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.
    - [PageContentCollection](../../docs-ref-autogen/onenote/onenote.pagecontentcollection.yml): A collection of PageContent objects, which represents the contents of a page.

- [Outline](../../docs-ref-autogen/onenote/onenote.outline.yml): A container for Paragraph objects. An Outline is a direct child of a PageContent object.

- [Image](../../docs-ref-autogen/onenote/onenote.image.yml): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.

- [Paragraph](../../docs-ref-autogen/onenote/onenote.paragraph.yml): A container for the visible content on a page. A Paragraph is a direct child of an Outline.
    - [ParagraphCollection](../../docs-ref-autogen/onenote/onenote.paragraphcollection.yml): A collection of Paragraph objects in an Outline.

- [RichText](../../docs-ref-autogen/onenote/onenote.richtext.yml): A RichText object.

- [Table](../../docs-ref-autogen/onenote/onenote.table.yml): A container for TableRow objects.

- [TableRow](../../docs-ref-autogen/onenote/onenote.tablerow.yml): A container for TableCell objects.
    - [TableRowCollection](../../docs-ref-autogen/onenote/onenote.tablerowcollection.yml): A collection of TableRow objects in a Table.
 
- [TableCell](../../docs-ref-autogen/onenote/onenote.tablecell.yml): A container for Paragraph objects.
    - [TableCellCollection](../../docs-ref-autogen/onenote/onenote.tablecellcollection.yml): A collection of TableCell objects in a TableRow.

## OneNote JavaScript API reference

For detailed information about OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](../../docs-ref-autogen/onenote.yml).

## See also

- [OneNote JavaScript API programming overview](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [Build your first OneNote add-in](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
