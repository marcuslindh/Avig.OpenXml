using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Avig.OpenXml.Tests;

/// <summary>Length, which Avig.Ord, Avig.Kalkyl and Avig.Bildspel share, and the PresentationML names, checked against the Open XML SDK.</summary>
public class LengthAndNamespacesTests
{
    [Fact]
    public void A_length_holds_every_unit_exactly_as_emu()
    {
        Assert.Equal(900_000, Length.Cm(2.5).Emu);
        Assert.Equal(25, Length.Cm(2.5).ToMm(), 9);
        Assert.Equal(139_700, Length.Pt(11).Emu);
        Assert.Equal(914_400, Length.Inch(1).Emu);
        Assert.Equal(9_525, Length.Px(1).Emu);
        Assert.Equal(Length.Pt(1), Length.FromTwips(20));
        Assert.Equal(22, Length.Pt(11).ToHalfPoints());
        Assert.Equal(Length.Pt(11), Length.FromHalfPoints(22));
        Assert.Equal(240, Length.Pt(12).ToTwips());
        Assert.Equal(Length.Cm(3), Length.Cm(1) + Length.Cm(2));
        Assert.True(Length.Mm(1) < Length.Cm(1));
        Assert.Equal("12pt", Length.Pt(12).ToString());
    }

    [Fact]
    public void Twips_and_half_points_past_what_a_length_holds_saturate_as_points_do()
    {
        // Before 0.3.2: 4.8e18 twips wrapped around to a negative length.
        Assert.Equal(long.MaxValue, Length.FromTwips(4_800_000_000_000_000_000).Emu);
        Assert.Equal(long.MaxValue, Length.FromTwips(long.MaxValue).Emu);
        Assert.Equal(long.MinValue, Length.FromTwips(long.MinValue).Emu);
        Assert.Equal(long.MaxValue, Length.FromTwips((long.MaxValue / 635) + 1).Emu);
        Assert.Equal((long.MaxValue / 635) * 635, Length.FromTwips(long.MaxValue / 635).Emu);
        Assert.Equal(long.MaxValue, Length.FromHalfPoints(long.MaxValue / 20).Emu);
        Assert.Equal(long.MinValue, Length.FromHalfPoints(long.MinValue).Emu);
        Assert.Equal((long.MaxValue / 6_350) * 6_350, Length.FromHalfPoints(long.MaxValue / 6_350).Emu);
        Assert.Equal(long.MaxValue, Length.Pt(1e300).Emu);
        Assert.Equal(long.MinValue, Length.Pt(-1e300).Emu);

        // Every value Word writes is as before.
        Assert.Equal(-635, Length.FromTwips(-1).Emu);
        Assert.Equal(1_440 * 635, Length.FromTwips(1_440).Emu);
        Assert.Equal(6_350, Length.FromHalfPoints(1).Emu);
        Assert.Equal(-6_350 * 3, Length.FromHalfPoints(-3).Emu);
        Assert.Equal(Length.Pt(1_638), Length.FromHalfPoints(3_276));
        Assert.Equal(Length.Inch(22), Length.FromTwips(31_680));
    }

    [Fact]
    public void The_presentation_content_types_and_relationships_are_the_sdks()
    {
        using MemoryStream stream = new();
        using PresentationDocument document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
        PresentationPart presentation = document.AddPresentationPart();
        SlidePart slide = presentation.AddNewPart<SlidePart>();
        SlideLayoutPart layout = slide.AddNewPart<SlideLayoutPart>();
        SlideMasterPart master = layout.AddNewPart<SlideMasterPart>();
        NotesSlidePart notes = slide.AddNewPart<NotesSlidePart>();
        NotesMasterPart notesMaster = presentation.AddNewPart<NotesMasterPart>();
        HandoutMasterPart handout = presentation.AddNewPart<HandoutMasterPart>();
        ThemePart theme = master.AddNewPart<ThemePart>();
        SlideCommentsPart comments = slide.AddNewPart<SlideCommentsPart>();
        CommentAuthorsPart authors = presentation.AddNewPart<CommentAuthorsPart>();
        PowerPointCommentPart modernComments = slide.AddNewPart<PowerPointCommentPart>();
        PowerPointAuthorsPart modernAuthors = presentation.AddNewPart<PowerPointAuthorsPart>();
        PresentationPropertiesPart properties = presentation.AddNewPart<PresentationPropertiesPart>();
        ViewPropertiesPart view = presentation.AddNewPart<ViewPropertiesPart>();
        TableStylesPart tableStyles = presentation.AddNewPart<TableStylesPart>();

        Assert.Equal(OpcNamespaces.PresentationContentType, presentation.ContentType);
        Assert.Equal(OpcNamespaces.SlideContentType, slide.ContentType);
        Assert.Equal(OpcNamespaces.SlideRelationship, slide.RelationshipType);
        Assert.Equal(OpcNamespaces.SlideLayoutContentType, layout.ContentType);
        Assert.Equal(OpcNamespaces.SlideLayoutRelationship, layout.RelationshipType);
        Assert.Equal(OpcNamespaces.SlideMasterContentType, master.ContentType);
        Assert.Equal(OpcNamespaces.SlideMasterRelationship, master.RelationshipType);
        Assert.Equal(OpcNamespaces.NotesSlideContentType, notes.ContentType);
        Assert.Equal(OpcNamespaces.NotesSlideRelationship, notes.RelationshipType);
        Assert.Equal(OpcNamespaces.NotesMasterContentType, notesMaster.ContentType);
        Assert.Equal(OpcNamespaces.NotesMasterRelationship, notesMaster.RelationshipType);
        Assert.Equal(OpcNamespaces.HandoutMasterContentType, handout.ContentType);
        Assert.Equal(OpcNamespaces.HandoutMasterRelationship, handout.RelationshipType);
        Assert.Equal(OpcNamespaces.ThemeContentType, theme.ContentType);
        Assert.Equal(OpcNamespaces.ThemeRelationship, theme.RelationshipType);
        Assert.Equal(OpcNamespaces.CommentsContentType, comments.ContentType);
        Assert.Equal(OpcNamespaces.CommentsRelationship, comments.RelationshipType);
        Assert.Equal(OpcNamespaces.CommentAuthorsContentType, authors.ContentType);
        Assert.Equal(OpcNamespaces.CommentAuthorsRelationship, authors.RelationshipType);
        Assert.Equal(OpcNamespaces.ModernCommentsContentType, modernComments.ContentType);
        Assert.Equal(OpcNamespaces.ModernCommentsRelationship, modernComments.RelationshipType);
        Assert.Equal(OpcNamespaces.ModernCommentAuthorsContentType, modernAuthors.ContentType);
        Assert.Equal(OpcNamespaces.ModernCommentAuthorsRelationship, modernAuthors.RelationshipType);
        Assert.Equal(OpcNamespaces.PresentationPropertiesContentType, properties.ContentType);
        Assert.Equal(OpcNamespaces.PresentationPropertiesRelationship, properties.RelationshipType);
        Assert.Equal(OpcNamespaces.ViewPropertiesContentType, view.ContentType);
        Assert.Equal(OpcNamespaces.ViewPropertiesRelationship, view.RelationshipType);
        Assert.Equal(OpcNamespaces.TableStylesContentType, tableStyles.ContentType);
        Assert.Equal(OpcNamespaces.TableStylesRelationship, tableStyles.RelationshipType);
        Assert.Equal(OpcNamespaces.PresentationMain, new DocumentFormat.OpenXml.Presentation.Slide().NamespaceUri);
        Assert.Equal(OpcNamespaces.DrawingMain, new DocumentFormat.OpenXml.Drawing.Blip().NamespaceUri);
        Assert.Equal(OpcNamespaces.DrawingChart, new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace().NamespaceUri);
    }

    [Theory]
    [InlineData(PresentationDocumentType.Presentation, OpcNamespaces.PresentationContentType)]
    [InlineData(PresentationDocumentType.MacroEnabledPresentation, OpcNamespaces.PresentationMacroEnabledContentType)]
    [InlineData(PresentationDocumentType.Template, OpcNamespaces.PresentationTemplateContentType)]
    [InlineData(PresentationDocumentType.MacroEnabledTemplate, OpcNamespaces.PresentationTemplateMacroEnabledContentType)]
    [InlineData(PresentationDocumentType.Slideshow, OpcNamespaces.PresentationSlideShowContentType)]
    [InlineData(PresentationDocumentType.MacroEnabledSlideshow, OpcNamespaces.PresentationSlideShowMacroEnabledContentType)]
    public void Each_kind_of_presentation_has_its_main_content_type(PresentationDocumentType kind, string contentType)
    {
        using MemoryStream stream = new();
        using PresentationDocument document = PresentationDocument.Create(stream, kind);

        Assert.Equal(contentType, document.AddPresentationPart().ContentType);
    }
}
