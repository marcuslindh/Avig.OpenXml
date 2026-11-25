namespace Avig.OpenXml;

/// <summary>Namespaces, relationship types and content types that every OPC package shares.</summary>
public static class OpcNamespaces
{
    /// <summary>The namespace of r:id and other relationship attributes in transitional OOXML.</summary>
    public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>The namespace of the elements in a .rels part.</summary>
    public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>The namespace of [Content_Types].xml.</summary>
    public const string ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>The namespace of Flat OPC (pkg:package), as Word writes it in "Word XML Document".</summary>
    public const string FlatOpc = "http://schemas.microsoft.com/office/2006/xmlPackage";

    /// <summary>The namespace of mc:Ignorable and mc:AlternateContent.</summary>
    public const string MarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>The relationship from the package to its main part.</summary>
    public const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    /// <summary>The relationship from the package to docProps/core.xml.</summary>
    public const string CorePropertiesRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    /// <summary>The relationship from the package to docProps/app.xml.</summary>
    public const string ExtendedPropertiesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

    /// <summary>The relationship from a part to an image.</summary>
    public const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    /// <summary>The relationship from a part to a URL (with TargetMode="External").</summary>
    public const string HyperlinkRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    /// <summary>The content type of a .rels part.</summary>
    public const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";

    /// <summary>The content type of docProps/core.xml.</summary>
    public const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";

    /// <summary>The content type of docProps/app.xml.</summary>
    public const string ExtendedPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

    /// <summary>The content type of a VBA project, which only macro-enabled files have.</summary>
    public const string VbaProjectContentType = "application/vnd.ms-office.vbaProject";

    /// <summary>The content type of an ActiveX control's properties (activeX1.xml), in Word, Excel and PowerPoint.</summary>
    public const string ActiveXContentType = "application/vnd.ms-office.activeX+xml";

    /// <summary>The content type of an ActiveX control's binary state (activeX1.bin).</summary>
    public const string ActiveXBinaryContentType = "application/vnd.ms-office.activeX";

    /// <summary>The content type of an embedded OLE object (oleObject1.bin), which can hold a program or a package of files.</summary>
    public const string OleObjectContentType = "application/vnd.openxmlformats-officedocument.oleObject";

    // --- DrawingML, shared by Word, Excel and PowerPoint ---

    /// <summary>The namespace of DrawingML's shapes, fills, lines, effects and text (a:).</summary>
    public const string DrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>The namespace of a chart part (c:chartSpace).</summary>
    public const string DrawingChart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>The namespace of the SVG extension of a:blip (asvg:svgBlip), from Office 2016.</summary>
    public const string DrawingSvg = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

    /// <summary>The relationship from a drawing, a slide or a graphic frame to a chart part.</summary>
    public const string ChartRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

    /// <summary>The content type of a chart part.</summary>
    public const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

    /// <summary>The relationship to a theme part.</summary>
    public const string ThemeRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    /// <summary>The content type of a theme part (theme1.xml), in Word, Excel and PowerPoint.</summary>
    public const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";

    // --- PresentationML ---

    /// <summary>The namespace of PresentationML (p:): the presentation, slides, layouts, masters and notes.</summary>
    public const string PresentationMain = "http://schemas.openxmlformats.org/presentationml/2006/main";

    /// <summary>The namespace of PowerPoint's modern (threaded) comments and their authors, from Microsoft 365 (p188:).</summary>
    public const string PresentationModernComments = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";

    /// <summary>The content type of presentation.xml in a .pptx.</summary>
    public const string PresentationContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    /// <summary>The content type of presentation.xml in a .pptm, a presentation with macros.</summary>
    public const string PresentationMacroEnabledContentType = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";

    /// <summary>The content type of presentation.xml in a .potx, a template.</summary>
    public const string PresentationTemplateContentType = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";

    /// <summary>The content type of presentation.xml in a .potm, a template with macros.</summary>
    public const string PresentationTemplateMacroEnabledContentType = "application/vnd.ms-powerpoint.template.macroEnabled.main+xml";

    /// <summary>The content type of presentation.xml in a .ppsx, a show that opens in slide show view.</summary>
    public const string PresentationSlideShowContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml";

    /// <summary>The content type of presentation.xml in a .ppsm, a show with macros.</summary>
    public const string PresentationSlideShowMacroEnabledContentType = "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml";

    /// <summary>The content type of a slide.</summary>
    public const string SlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

    /// <summary>The content type of a slide layout.</summary>
    public const string SlideLayoutContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";

    /// <summary>The content type of a slide master.</summary>
    public const string SlideMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

    /// <summary>The content type of a slide's speaker notes.</summary>
    public const string NotesSlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

    /// <summary>The content type of the notes master.</summary>
    public const string NotesMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";

    /// <summary>The content type of the handout master.</summary>
    public const string HandoutMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";

    /// <summary>The content type of a slide's classic comments (comment1.xml).</summary>
    public const string CommentsContentType = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";

    /// <summary>The content type of the authors of classic comments (commentAuthors.xml).</summary>
    public const string CommentAuthorsContentType = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";

    /// <summary>The content type of a slide's modern comments (modernComment_*.xml).</summary>
    public const string ModernCommentsContentType = "application/vnd.ms-powerpoint.comments+xml";

    /// <summary>The content type of the authors of modern comments (authors.xml).</summary>
    public const string ModernCommentAuthorsContentType = "application/vnd.ms-powerpoint.authors+xml";

    /// <summary>The content type of presProps.xml: the presentation's show settings.</summary>
    public const string PresentationPropertiesContentType = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";

    /// <summary>The content type of viewProps.xml: how PowerPoint last showed the presentation.</summary>
    public const string ViewPropertiesContentType = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";

    /// <summary>The content type of tableStyles.xml.</summary>
    public const string TableStylesContentType = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";

    /// <summary>The relationship from presentation.xml to a slide.</summary>
    public const string SlideRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    /// <summary>The relationship from a slide to its layout.</summary>
    public const string SlideLayoutRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

    /// <summary>The relationship from presentation.xml or a layout to a slide master.</summary>
    public const string SlideMasterRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    /// <summary>The relationship from a slide to its speaker notes.</summary>
    public const string NotesSlideRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

    /// <summary>The relationship from presentation.xml or a notes slide to the notes master.</summary>
    public const string NotesMasterRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";

    /// <summary>The relationship from presentation.xml to the handout master.</summary>
    public const string HandoutMasterRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";

    /// <summary>The relationship from a slide to its classic comments.</summary>
    public const string CommentsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    /// <summary>The relationship from presentation.xml to the authors of classic comments.</summary>
    public const string CommentAuthorsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";

    /// <summary>The relationship from a slide to its modern comments.</summary>
    public const string ModernCommentsRelationship = "http://schemas.microsoft.com/office/2018/10/relationships/comments";

    /// <summary>The relationship from presentation.xml to the authors of modern comments.</summary>
    public const string ModernCommentAuthorsRelationship = "http://schemas.microsoft.com/office/2018/10/relationships/authors";

    /// <summary>The relationship from presentation.xml to presProps.xml.</summary>
    public const string PresentationPropertiesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";

    /// <summary>The relationship from presentation.xml to viewProps.xml.</summary>
    public const string ViewPropertiesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";

    /// <summary>The relationship from presentation.xml to tableStyles.xml.</summary>
    public const string TableStylesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
}
