# OneNote add-ins JavaScript API reference

Applies to: OneNote Online

The following links show the high level OneNote objects available in the API. Each object page link contains a description of the properties, relationships, and methods available on the object. Explore these links to learn more. 
	
- [Application](../../api/onenote/onenote.application.yml): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.

- [Notebook](../../api/onenote/onenote.notebook.yml): A notebook. Notebooks contain section groups and sections.

   - [NotebookCollection](../../api/onenote/onenote.notebookcollection.yml): A collection of notebooks.

- [SectionGroup](../../api/onenote/onenote.sectiongroup.yml): A section group. Section groups contain section groups and sections.

   - [SectionGroupCollection](../../api/onenote/onenote.sectiongroupcollection.yml): A collection of section groups.

- [Section](../../api/onenote/onenote.section.yml): A section. Sections contain pages.

   - [SectionCollection](../../api/onenote/onenote.sectioncollection.yml): A collection of sections.

- [Page](../../api/onenote/onenote.page.yml): A page. Pages contain PageContent objects.

   - [PageCollection](../../api/onenote/onenote.pagecollection.yml): A collection of pages.

- [PageContent](../../api/onenote/onenote.pagecontent.yml): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.

   - [PageContentCollection](../../api/onenote/onenote.pagecontentcollection.yml): A collection of PageContent objects, which represents the contents of a page.

- [Outline](../../api/onenote/onenote.outline.yml): A container for Paragraph objects. An Outline is a direct child of a PageContent object.

- [Image](../../api/onenote/onenote.image.yml): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.

- [Paragraph](../../api/onenote/onenote.paragraph.yml): A container for the visible content on a page. A Paragraph is a direct child of an Outline.

  - [ParagraphCollection](../../api/onenote/onenote.paragraphcollection.yml): A collection of Paragraph objects in an Outline.

- [RichText](../../api/onenote/onenote.richtext.yml): A RichText object.

- [Table](../../api/onenote/onenote.table.yml): A container for TableRow objects.

- [TableRow](../../api/onenote/onenote.tablerow.yml): A container for TableCell objects.

  - [TableRowCollection](../../api/onenote/onenote.tablerowcollection.yml): A collection of TableRow objects in a Table.
 
- [TableCell](../../api/onenote/onenote.tablecell.yml): A container for Paragraph objects.

  - [TableCellCollection](../../api/onenote/onenote.tablecellcollection.yml): A collection of TableCell objects in a TableRow.
		
## Additional resources

- [OneNote JavaScript API programming overview](https://docs.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [Build your first OneNote add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
