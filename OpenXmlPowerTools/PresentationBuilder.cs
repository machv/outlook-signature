/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2013.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

***************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class SlideSource
    {
        public PmlDocument PmlDocument { get; set; }
        public int Start { get; set; }
        public int Count { get; set; }
        public bool KeepMaster { get; set; }

        public SlideSource(PmlDocument source, bool keepMaster)
        {
            PmlDocument = source;
            Start = 0;
            Count = Int32.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(PmlDocument source, int start, bool keepMaster)
        {
            PmlDocument = source;
            Start = start;
            Count = Int32.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(PmlDocument source, int start, int count, bool keepMaster)
        {
            PmlDocument = source;
            Start = start;
            Count = count;
            KeepMaster = keepMaster;
        }
    }

    public static class PresentationBuilder
    {
        public static void BuildPresentation(List<SlideSource> sources, string fileName)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument())
            {
                using (PresentationDocument output = streamDoc.GetPresentationDocument())
                {
                    BuildPresentation(sources, output);
                    output.Dispose();
                }
                streamDoc.GetModifiedDocument().SaveAs(fileName);
            }
        }

        public static WmlDocument BuildPresentation(List<SlideSource> sources)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument())
            {
                using (PresentationDocument output = streamDoc.GetPresentationDocument())
                {
                    BuildPresentation(sources, output);
                    output.Dispose();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        private static void BuildPresentation(List<SlideSource> sources, PresentationDocument output)
        {
            if (RelationshipMarkup == null)
                RelationshipMarkup = new Dictionary<XName, XName[]>()
                {
                    { A.audioFile,        new [] { R.link }},
                    { A.videoFile,        new [] { R.link }},
                    { A.wavAudioFile,     new [] { R.embed }},
                    { A.blip,             new [] { R.embed, R.link }},
                    { A.hlinkClick,       new [] { R.id }},
                    { A.relIds,           new [] { R.cs, R.dm, R.lo, R.qs }},
                    { C.chart,            new [] { R.id }},
                    { C.externalData,     new [] { R.id }},
                    { C.userShapes,       new [] { R.id }},
                    { DGM.relIds,         new [] { R.cs, R.dm, R.lo, R.qs }},
                    { P.oleObj,           new [] { R.id }},
                    { P.snd,              new [] { R.embed }},
                    { VML.fill,           new [] { R.id }},
                    { VML.imagedata,      new [] { R.href, R.id, R.pict, O.relid }},
                    { VML.stroke,         new [] { R.id }},
                    { WNE.toolbarData,    new [] { R.id }},
                };

            List<ImageData> images = new List<ImageData>();
            XDocument mainPart = output.PresentationPart.GetXDocument();
            mainPart.Declaration.Standalone = "yes";
            mainPart.Declaration.Encoding = "UTF-8";
            output.PresentationPart.PutXDocument();
            int sourceNum = 0;
            SlideMasterPart currentMasterPart = null;
            foreach (SlideSource source in sources)
            {
                using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(source.PmlDocument))
                using (PresentationDocument doc = streamDoc.GetPresentationDocument())
                {
                    try
                    {
                        if (sourceNum == 0)
                            CopyPresentationParts(doc, output, images);
                        currentMasterPart = AppendSlides(doc, output, source.Start, source.Count, source.KeepMaster, images, currentMasterPart);
                    }
                    catch (PresentationBuilderInternalException dbie)
                    {
                        if (dbie.Message.Contains("{0}"))
                            throw new PresentationBuilderException(string.Format(dbie.Message, sourceNum));
                        else
                            throw dbie;
                    }
                }
                sourceNum++;
            }
            foreach (var part in output.GetAllParts())
                if (part.Annotation<XDocument>() != null)
                    part.PutXDocument();
        }

        // Copy handout master, notes master, presentation properties and view properties, if they exist
        private static void CopyPresentationParts(PresentationDocument sourceDocument, PresentationDocument newDocument, List<ImageData> images)
        {
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();

            // Copy slide and note slide sizes
            XDocument oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();
            XElement oldElement = oldPresentationDoc.Root.Elements(P.sldSz).FirstOrDefault();
            if (oldElement != null)
                newPresentation.Root.Element(P.notesSz).AddBeforeSelf(oldElement);

            // Copy Handout Master
            if (sourceDocument.PresentationPart.HandoutMasterPart != null)
            {
                HandoutMasterPart oldMaster = sourceDocument.PresentationPart.HandoutMasterPart;
                HandoutMasterPart newMaster = newDocument.PresentationPart.AddNewPart<HandoutMasterPart>();

                // Copy theme for master
                ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
                newThemePart.PutXDocument(oldMaster.ThemePart.GetXDocument());
                CopyRelatedPartsForContentParts(newDocument, oldMaster.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images);

                // Copy master
                newMaster.PutXDocument(oldMaster.GetXDocument());
                AddRelationships(oldMaster, newMaster, new[] { newMaster.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldMaster, newMaster, new[] { newMaster.GetXDocument().Root }, images);

                newPresentation.Root.Element(P.sldMasterIdLst).AddAfterSelf(
                    new XElement(P.handoutMasterIdLst, new XElement(P.handoutMasterId,
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster)))));
            }

            // Copy Notes Master
            CopyNotesMaster(sourceDocument, newDocument, images);

            // Copy Presentation Properties
            if (sourceDocument.PresentationPart.PresentationPropertiesPart != null)
            {
                PresentationPropertiesPart newPart = newDocument.PresentationPart.AddNewPart<PresentationPropertiesPart>();
                newPart.PutXDocument(sourceDocument.PresentationPart.PresentationPropertiesPart.GetXDocument());
            }

            // Copy View Properties
            if (sourceDocument.PresentationPart.ViewPropertiesPart != null)
            {
                ViewPropertiesPart newPart = newDocument.PresentationPart.AddNewPart<ViewPropertiesPart>();
                newPart.PutXDocument(sourceDocument.PresentationPart.ViewPropertiesPart.GetXDocument());
            }
        }

        private static SlideMasterPart AppendSlides(PresentationDocument sourceDocument, PresentationDocument newDocument,
            int start, int count, bool keepMaster, List<ImageData> images, SlideMasterPart currentMasterPart)
        {
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();
            if (newPresentation.Root.Element(P.sldIdLst) == null)
                newPresentation.Root.Add(new XElement(P.sldIdLst));
            uint newID = 256;
            var ids = newPresentation.Root.Descendants(P.sldId).Select(f => (uint)f.Attribute(NoNamespace.id));
            if (ids.Any())
                newID = ids.Max() + 1;
            var slideList = sourceDocument.PresentationPart.GetXDocument().Root.Descendants(P.sldId);
            while (count > 0 && start < slideList.Count())
            {
                SlidePart slide = (SlidePart)sourceDocument.PresentationPart.GetPartById(slideList.ElementAt(start).Attribute(R.id).Value);
                if (currentMasterPart == null || keepMaster)
                    currentMasterPart = CopyMasterSlide(sourceDocument, slide.SlideLayoutPart.SlideMasterPart, newDocument, newPresentation, images);
                SlidePart newSlide = newDocument.PresentationPart.AddNewPart<SlidePart>();
                newSlide.PutXDocument(slide.GetXDocument());
                AddRelationships(slide, newSlide, new[] { newSlide.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, slide, newSlide, new[] { newSlide.GetXDocument().Root }, images);
                CopyTableStyles(sourceDocument, newDocument, slide, newSlide);
                if (slide.NotesSlidePart != null)
                {
                    if (newDocument.PresentationPart.NotesMasterPart == null)
                        CopyNotesMaster(sourceDocument, newDocument, images);
                    NotesSlidePart newPart = newSlide.AddNewPart<NotesSlidePart>();
                    newPart.PutXDocument(slide.NotesSlidePart.GetXDocument());
                    newPart.AddPart(newSlide);
                    newPart.AddPart(newDocument.PresentationPart.NotesMasterPart);
                }

                string layoutName = slide.SlideLayoutPart.GetXDocument().Root.Element(P.cSld).Attribute(NoNamespace.name).Value;
                foreach (SlideLayoutPart layoutPart in currentMasterPart.SlideLayoutParts)
                    if (layoutPart.GetXDocument().Root.Element(P.cSld).Attribute(NoNamespace.name).Value == layoutName)
                    {
                        newSlide.AddPart(layoutPart);
                        break;
                    }
                if (newSlide.SlideLayoutPart == null)
                    newSlide.AddPart(currentMasterPart.SlideLayoutParts.First());  // Cannot find matching layout part

                if (slide.SlideCommentsPart != null)
                    CopyComments(sourceDocument, newDocument, slide, newSlide);

                newPresentation.Root.Element(P.sldIdLst).Add(new XElement(P.sldId,
                    new XAttribute(NoNamespace.id, newID.ToString()),
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newSlide))));
                newID++;
                start++;
                count--;
            }
            return currentMasterPart;
        }

        private static SlideMasterPart CopyMasterSlide(PresentationDocument sourceDocument, SlideMasterPart sourceMasterPart,
            PresentationDocument newDocument, XDocument newPresentation, List<ImageData> images)
        {
            // Search for existing master slide with same theme name
            XDocument oldTheme = sourceMasterPart.ThemePart.GetXDocument();
            String themeName = oldTheme.Root.Attribute(NoNamespace.name).Value;
            foreach (SlideMasterPart master in newDocument.PresentationPart.GetPartsOfType<SlideMasterPart>())
            {
                XDocument themeDoc = master.ThemePart.GetXDocument();
                if (themeDoc.Root.Attribute(NoNamespace.name).Value == themeName)
                    return master;
            }

            SlideMasterPart newMaster = newDocument.PresentationPart.AddNewPart<SlideMasterPart>();
            XDocument sourceMaster = sourceMasterPart.GetXDocument();

            // Add to presentation slide master list, need newID for layout IDs also
            uint newID = 2147483648;
            var ids = newPresentation.Root.Descendants(P.sldMasterId).Select(f => (uint)f.Attribute(NoNamespace.id));
            if (ids.Any())
            {
                newID = ids.Max();
                XElement maxMaster = newPresentation.Root.Descendants(P.sldMasterId).Where(f => (uint)f.Attribute(NoNamespace.id) == newID).FirstOrDefault();
                SlideMasterPart maxMasterPart = (SlideMasterPart)newDocument.PresentationPart.GetPartById(maxMaster.Attribute(R.id).Value);
                newID += (uint)maxMasterPart.GetXDocument().Root.Descendants(P.sldLayoutId).Count() + 1;
            }
            newPresentation.Root.Element(P.sldMasterIdLst).Add(new XElement(P.sldMasterId,
                new XAttribute(NoNamespace.id, newID.ToString()),
                new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster))));
            newID++;

            ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
            if (newDocument.PresentationPart.ThemePart == null)
                newThemePart = newDocument.PresentationPart.AddPart(newThemePart);
            newThemePart.PutXDocument(oldTheme);
            CopyRelatedPartsForContentParts(newDocument, sourceMasterPart.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images);
            foreach (SlideLayoutPart layoutPart in sourceMasterPart.SlideLayoutParts)
            {
                SlideLayoutPart newLayout = newMaster.AddNewPart<SlideLayoutPart>();
                newLayout.PutXDocument(layoutPart.GetXDocument());
                AddRelationships(layoutPart, newLayout, new[] { newLayout.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, layoutPart, newLayout, new[] { newLayout.GetXDocument().Root }, images);
                newLayout.AddPart(newMaster);
                string resID = sourceMasterPart.GetIdOfPart(layoutPart);
                XElement entry = sourceMaster.Root.Descendants(P.sldLayoutId).Where(f => f.Attribute(R.id).Value == resID).FirstOrDefault();
                entry.Attribute(R.id).SetValue(newMaster.GetIdOfPart(newLayout));
                entry.Attribute(NoNamespace.id).SetValue(newID.ToString());
                newID++;
            }
            newMaster.PutXDocument(sourceMaster);
            AddRelationships(sourceMasterPart, newMaster, new[] { newMaster.GetXDocument().Root });
            CopyRelatedPartsForContentParts(newDocument, sourceMasterPart, newMaster, new[] { newMaster.GetXDocument().Root }, images);

            return newMaster;
        }

        // Copies notes master and notesSz element from presentation
        private static void CopyNotesMaster(PresentationDocument sourceDocument, PresentationDocument newDocument, List<ImageData> images)
        {
            // Copy notesSz element from presentation
            XDocument newPresentation = newDocument.PresentationPart.GetXDocument();
            XDocument oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();
            XElement oldElement = oldPresentationDoc.Root.Element(P.notesSz);
            newPresentation.Root.Element(P.notesSz).ReplaceWith(oldElement);

            // Copy Notes Master
            if (sourceDocument.PresentationPart.NotesMasterPart != null)
            {
                NotesMasterPart oldMaster = sourceDocument.PresentationPart.NotesMasterPart;
                NotesMasterPart newMaster = newDocument.PresentationPart.AddNewPart<NotesMasterPart>();

                // Copy theme for master
                ThemePart newThemePart = newMaster.AddNewPart<ThemePart>();
                newThemePart.PutXDocument(oldMaster.ThemePart.GetXDocument());
                CopyRelatedPartsForContentParts(newDocument, oldMaster.ThemePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images);

                // Copy master
                newMaster.PutXDocument(oldMaster.GetXDocument());
                AddRelationships(oldMaster, newMaster, new[] { newMaster.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldMaster, newMaster, new[] { newMaster.GetXDocument().Root }, images);

                newPresentation.Root.Element(P.sldMasterIdLst).AddAfterSelf(
                    new XElement(P.notesMasterIdLst, new XElement(P.notesMasterId,
                    new XAttribute(R.id, newDocument.PresentationPart.GetIdOfPart(newMaster)))));
            }
        }

        private static void CopyComments(PresentationDocument oldDocument, PresentationDocument newDocument, SlidePart oldSlide, SlidePart newSlide)
        {
            newSlide.AddNewPart<SlideCommentsPart>();
            newSlide.SlideCommentsPart.PutXDocument(oldSlide.SlideCommentsPart.GetXDocument());
            XDocument newSlideComments = newSlide.SlideCommentsPart.GetXDocument();
            XDocument oldAuthors = oldDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
            foreach (XElement comment in newSlideComments.Root.Elements(P.cm))
            {
                XElement newAuthor = FindCommentsAuthor(newDocument, comment, oldAuthors);
                // Update last index value for new comment
                comment.Attribute(NoNamespace.authorId).SetValue(newAuthor.Attribute(NoNamespace.id).Value);
                uint lastIndex = Convert.ToUInt32(newAuthor.Attribute(NoNamespace.lastIdx).Value);
                comment.Attribute(NoNamespace.idx).SetValue(lastIndex.ToString());
                newAuthor.Attribute(NoNamespace.lastIdx).SetValue(Convert.ToString(lastIndex + 1));
            }
        }

        private static XElement FindCommentsAuthor(PresentationDocument newDocument, XElement comment, XDocument oldAuthors)
        {
            XElement oldAuthor = oldAuthors.Root.Elements(P.cmAuthor).Where(
                f => f.Attribute(NoNamespace.id).Value == comment.Attribute(NoNamespace.authorId).Value).FirstOrDefault();
            XElement newAuthor = null;
            if (newDocument.PresentationPart.CommentAuthorsPart == null)
            {
                newDocument.PresentationPart.AddNewPart<CommentAuthorsPart>();
                newDocument.PresentationPart.CommentAuthorsPart.PutXDocument(new XDocument(new XElement(P.cmAuthorLst,
                    new XAttribute(XNamespace.Xmlns + "a", A.a),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XAttribute(XNamespace.Xmlns + "p", P.p))));
            }
            XDocument authors = newDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
            newAuthor = authors.Root.Elements(P.cmAuthor).Where(
                f => f.Attribute(NoNamespace.initials).Value == oldAuthor.Attribute(NoNamespace.initials).Value).FirstOrDefault();
            if (newAuthor == null)
            {
                uint newID = 0;
                var ids = authors.Root.Descendants(P.cmAuthor).Select(f => (uint)f.Attribute(NoNamespace.id));
                if (ids.Any())
                    newID = ids.Max() + 1;

                newAuthor = new XElement(P.cmAuthor, new XAttribute(NoNamespace.id, newID.ToString()),
                    new XAttribute(NoNamespace.name, oldAuthor.Attribute(NoNamespace.name).Value),
                    new XAttribute(NoNamespace.initials, oldAuthor.Attribute(NoNamespace.initials).Value),
                    new XAttribute(NoNamespace.lastIdx, "1"), new XAttribute(NoNamespace.clrIdx, newID.ToString()));
                authors.Root.Add(newAuthor);
            }

            return newAuthor;
        }

        private static void CopyTableStyles(PresentationDocument oldDocument, PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart)
        {
            foreach (XElement table in newContentPart.GetXDocument().Descendants(A.tableStyleId))
            {
                string styleId = table.Value;
                if (string.IsNullOrEmpty(styleId))
                    continue;

                // Find old style
                if (oldDocument.PresentationPart.TableStylesPart == null)
                    continue;
                XDocument oldTableStyles = oldDocument.PresentationPart.TableStylesPart.GetXDocument();
                XElement oldStyle = oldTableStyles.Root.Elements(A.tblStyle).Where(f => f.Attribute(NoNamespace.styleId).Value == styleId).FirstOrDefault();
                if (oldStyle == null)
                    continue;

                // Create new TableStylesPart, if needed
                XDocument tableStyles = null;
                if (newDocument.PresentationPart.TableStylesPart == null)
                {
                    TableStylesPart newStylesPart = newDocument.PresentationPart.AddNewPart<TableStylesPart>();
                    tableStyles = new XDocument(new XElement(A.tblStyleLst,
                        new XAttribute(XNamespace.Xmlns + "a", A.a),
                        new XAttribute(NoNamespace.def, styleId)));
                    newStylesPart.PutXDocument(tableStyles);
                }
                else
                    tableStyles = newDocument.PresentationPart.TableStylesPart.GetXDocument();

                // Search new TableStylesPart to see if it contains the ID
                if (tableStyles.Root.Elements(A.tblStyle).Where(f => f.Attribute(NoNamespace.styleId).Value == styleId).FirstOrDefault() != null)
                    continue;

                // Copy style to new part
                tableStyles.Root.Add(oldStyle);
            }

        }

        private static void CopyRelatedPartsForContentParts(PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == VML.imagedata || d.Name == VML.fill || d.Name == VML.stroke || d.Name == A.blip)
                .ToList();
            foreach (XElement imageReference in relevantElements)
            {
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.embed, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.pict, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.id, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, O.relid, images);
            }

            foreach (XElement diagramReference in newContent.DescendantsAndSelf().Where(d => d.Name == DGM.relIds || d.Name == A.relIds))
            {
                // dm attribute
                string relId = diagramReference.Attribute(R.dm).Value;
                try
                {
                    OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                    continue;
                }
                catch (ArgumentOutOfRangeException)
                {
                    try
                    {
                        ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                        continue;
                    }
                    catch (KeyNotFoundException)
                    {
                    }
                }
                OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                OpenXmlPart newPart = newContentPart.AddNewPart<DiagramDataPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.dm).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // lo attribute
                relId = diagramReference.Attribute(R.lo).Value;
                try
                {
                    OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                    continue;
                }
                catch (ArgumentOutOfRangeException)
                {
                    try
                    {
                        ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                        continue;
                    }
                    catch (KeyNotFoundException)
                    {
                    }
                }
                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramLayoutDefinitionPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.lo).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // qs attribute
                relId = diagramReference.Attribute(R.qs).Value;
                try
                {
                    OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                    continue;
                }
                catch (ArgumentOutOfRangeException)
                {
                    try
                    {
                        ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                        continue;
                    }
                    catch (KeyNotFoundException)
                    {
                    }
                }
                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramStylePart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.qs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // cs attribute
                relId = diagramReference.Attribute(R.cs).Value;
                try
                {
                    OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                    continue;
                }
                catch (ArgumentOutOfRangeException)
                {
                    try
                    {
                        ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                        continue;
                    }
                    catch (KeyNotFoundException)
                    {
                    }
                }
                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramColorsPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.cs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);
            }

            foreach (XElement oleReference in newContent.DescendantsAndSelf(P.oleObj))
            {
                string relId = oleReference.Attribute(R.id).Value;
                try
                {
                    // First look to see if this relId has already been added to the new document.
                    // This is necessary for those parts that get processed with both old and new ids, such as the comments
                    // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                    // in that case.
                    try
                    {
                        OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                        continue;
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        try
                        {
                            ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                            continue;
                        }
                        catch (KeyNotFoundException)
                        {
                            // nothing to do
                        }
                    }

                    OpenXmlPart oldPart = oldContentPart.GetPartById(relId);
                    OpenXmlPart newPart = null;
                    if (oldPart is EmbeddedObjectPart)
                    {
                        if (newContentPart is SlidePart)
                            newPart = ((SlidePart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                        if (newContentPart is SlideMasterPart)
                            newPart = ((SlideMasterPart)newContentPart).AddEmbeddedObjectPart(oldPart.ContentType);
                    }
                    else if (oldPart is EmbeddedPackagePart)
                    {
                        if (newContentPart is SlidePart)
                            newPart = ((SlidePart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is SlideMasterPart)
                            newPart = ((SlideMasterPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                        if (newContentPart is ChartPart)
                            newPart = ((ChartPart)newContentPart).AddEmbeddedPackagePart(oldPart.ContentType);
                    }
                    using (Stream oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                    using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                    {
                        int byteCount;
                        byte[] buffer = new byte[65536];
                        while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                            newObject.Write(buffer, 0, byteCount);
                    }
                    oleReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                }
                catch (ArgumentOutOfRangeException)
                {
                    ExternalRelationship er = oldContentPart.GetExternalRelationship(relId);
                    ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    oleReference.Attribute(R.id).Value = newEr.Id;
                }
            }

            foreach (XElement chartReference in newContent.DescendantsAndSelf(C.chart))
            {
                try
                {
                    string relId = (string)chartReference.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    try
                    {
                        OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                        continue;
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        try
                        {
                            ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                            continue;
                        }
                        catch (KeyNotFoundException)
                        {
                        }
                    }
                    ChartPart oldPart = (ChartPart)oldContentPart.GetPartById(relId);
                    XDocument oldChart = oldPart.GetXDocument();
                    ChartPart newPart = newContentPart.AddNewPart<ChartPart>();
                    XDocument newChart = newPart.GetXDocument();
                    newChart.Add(oldChart.Root);
                    chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                    CopyChartObjects(oldPart, newPart);
                    CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newChart.Root }, images);
                }
                catch (ArgumentOutOfRangeException)
                {
                    continue;
                }
            }

            foreach (XElement userShape in newContent.DescendantsAndSelf(C.userShapes))
            {
                try
                {
                    string relId = (string)userShape.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    try
                    {
                        OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                        continue;
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        try
                        {
                            ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                            continue;
                        }
                        catch (KeyNotFoundException)
                        {
                        }
                    }
                    ChartDrawingPart oldPart = (ChartDrawingPart)oldContentPart.GetPartById(relId);
                    XDocument oldXDoc = oldPart.GetXDocument();
                    ChartDrawingPart newPart = newContentPart.AddNewPart<ChartDrawingPart>();
                    XDocument newXDoc = newPart.GetXDocument();
                    newXDoc.Add(oldXDoc.Root);
                    userShape.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                    AddRelationships(oldPart, newPart, newContent);
                    CopyRelatedPartsForContentParts(newDocument, oldPart, newPart, new[] { newXDoc.Root }, images);
                }
                catch (ArgumentOutOfRangeException)
                {
                    continue;
                }
            }

            foreach (XElement tags in newContent.DescendantsAndSelf(P.tags))
            {
                try
                {
                    string relId = (string)tags.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    try
                    {
                        OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                        continue;
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        try
                        {
                            ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                            continue;
                        }
                        catch (KeyNotFoundException)
                        {
                        }
                    }
                    UserDefinedTagsPart oldPart = (UserDefinedTagsPart)oldContentPart.GetPartById(relId);
                    XDocument oldXDoc = oldPart.GetXDocument();
                    UserDefinedTagsPart newPart = newContentPart.AddNewPart<UserDefinedTagsPart>();
                    XDocument newXDoc = newPart.GetXDocument();
                    newXDoc.Add(oldXDoc.Root);
                    tags.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                }
                catch (ArgumentOutOfRangeException)
                {
                    continue;
                }
            }

            foreach (XElement soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == A.wavAudioFile))
                CopyRelatedSound(newDocument, oldContentPart, newContentPart, soundReference, R.embed);

            if (oldContentPart is SlidePart && newContentPart is SlidePart)
            {
                foreach (XElement soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == P.snd))
                    CopyRelatedSound(newDocument, oldContentPart, newContentPart, soundReference, R.embed);

                // Transitional: Copy VML Drawing parts, implicit relationship
                foreach (VmlDrawingPart vmlPart in ((SlidePart)oldContentPart).VmlDrawingParts)
                {
                    VmlDrawingPart newVmlPart = ((SlidePart)newContentPart).AddNewPart<VmlDrawingPart>();
                    newVmlPart.PutXDocument(vmlPart.GetXDocument());
                    AddRelationships(vmlPart, newVmlPart, new[] { newVmlPart.GetXDocument().Root });
                    CopyRelatedPartsForContentParts(newDocument, vmlPart, newVmlPart, new[] { newVmlPart.GetXDocument().Root }, images);
                }
            }
            if (oldContentPart is SlideMasterPart && newContentPart is SlideMasterPart)
            {
                foreach (XElement soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == P.snd))
                    CopyRelatedSound(newDocument, oldContentPart, newContentPart, soundReference, R.embed);

                // Transitional: Copy VML Drawing parts, implicit relationship
                foreach (VmlDrawingPart vmlPart in ((SlideMasterPart)oldContentPart).VmlDrawingParts)
                {
                    VmlDrawingPart newVmlPart = ((SlideMasterPart)newContentPart).AddNewPart<VmlDrawingPart>();
                    newVmlPart.PutXDocument(vmlPart.GetXDocument());
                    AddRelationships(vmlPart, newVmlPart, new[] { newVmlPart.GetXDocument().Root });
                    CopyRelatedPartsForContentParts(newDocument, vmlPart, newVmlPart, new[] { newVmlPart.GetXDocument().Root }, images);
                }
            }
        }

        private static void CopyChartObjects(ChartPart oldChart, ChartPart newChart)
        {
            foreach (XElement dataReference in newChart.GetXDocument().Descendants(C.externalData))
            {
                string relId = dataReference.Attribute(R.id).Value;
                try
                {
                    EmbeddedPackagePart oldPart = (EmbeddedPackagePart)oldChart.GetPartById(relId);
                    EmbeddedPackagePart newPart = newChart.AddEmbeddedPackagePart(oldPart.ContentType);
                    using (Stream oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                    using (Stream newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                    {
                        int byteCount;
                        byte[] buffer = new byte[65536];
                        while ((byteCount = oldObject.Read(buffer, 0, 65536)) != 0)
                            newObject.Write(buffer, 0, byteCount);
                    }
                    dataReference.Attribute(R.id).Value = newChart.GetIdOfPart(newPart);
                }
                catch (ArgumentOutOfRangeException)
                {
                    ExternalRelationship oldRelationship = oldChart.GetExternalRelationship(relId);
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldRel = oldChart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new PresentationBuilderInternalException("Internal Error 0007");
                    newChart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    dataReference.Attribute(R.id).Value = newRid;
                }
            }
        }

        private static Dictionary<XName, XName[]> RelationshipMarkup = null;

        private static void UpdateContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid, string newRid)
        {
            foreach (var attributeName in RelationshipMarkup[elementToModify])
            {
                var elementsToUpdate = newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid);
                foreach (var element in elementsToUpdate)
                    element.Attribute(attributeName).Value = newRid;
            }
        }

        private static void RemoveContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid)
        {
            foreach (var attributeName in RelationshipMarkup[elementToModify])
            {
                newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid).Remove();
            }
        }

        private static void AddRelationships(OpenXmlPart oldPart, OpenXmlPart newPart, IEnumerable<XElement> newContent)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => RelationshipMarkup.ContainsKey(d.Name) &&
                    d.Attributes().Any(a => RelationshipMarkup[d.Name].Contains(a.Name)))
                .ToList();
            foreach (var e in relevantElements)
            {
                if (e.Name == A.hlinkClick)
                {
                    string relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null) {
                        //TODO Issue with reference to another part: var temp = oldPart.GetPartById(relId);
                        RemoveContent(newContent, e.Name, relId);
                        continue;
                    }
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == VML.imagedata)
                {
                    string relId = (string)e.Attribute(R.href);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new PresentationBuilderInternalException("Internal Error 0006");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == A.blip || e.Name == A.audioFile || e.Name == A.videoFile)
                {
                    string relId = (string)e.Attribute(R.link);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    Guid g = Guid.NewGuid();
                    string newRid = "R" + g.ToString().Replace("-", "");
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        continue;
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
            }
        }

        private static void CopyRelatedImage(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName,
            List<ImageData> images)
        {
            string relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;
            try
            {
                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                try
                {
                    OpenXmlPart tempPart = newContentPart.GetPartById(relId);
                    return;
                }
                catch (ArgumentOutOfRangeException)
                {
                    try
                    {
                        ExternalRelationship tempEr = newContentPart.GetExternalRelationship(relId);
                        return;
                    }
                    catch (KeyNotFoundException)
                    {
                        // nothing to do
                    }
                }

                ImagePart oldPart = (ImagePart)oldContentPart.GetPartById(relId);
                ImageData temp = ManageImageCopy(oldPart, newContentPart, images);
                if (temp.ResourceID == null)
                {
                    ImagePart newPart = null;
                    if (newContentPart is ThemePart)
                        newPart = ((ThemePart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is DocumentSettingsPart)
                        newPart = ((DocumentSettingsPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ChartPart)
                        newPart = ((ChartPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is DiagramDataPart)
                        newPart = ((DiagramDataPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is ChartDrawingPart)
                        newPart = ((ChartDrawingPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is SlidePart)
                        newPart = ((SlidePart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is SlideMasterPart)
                        newPart = ((SlideMasterPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is SlideLayoutPart)
                        newPart = ((SlideLayoutPart)newContentPart).AddImagePart(oldPart.ContentType);
                    if (newContentPart is VmlDrawingPart)
                        newPart = ((VmlDrawingPart)newContentPart).AddImagePart(oldPart.ContentType);
                    temp.ResourceID = newContentPart.GetIdOfPart(newPart);
                    temp.WriteImage(newPart);
                    imageReference.Attribute(attributeName).Value = temp.ResourceID;
                }
                else
                {
                    try
                    {
                        newContentPart.GetReferenceRelationship(temp.ResourceID);
                        imageReference.Attribute(attributeName).Value = temp.ResourceID;
                    }
                    catch (KeyNotFoundException)
                    {
                        ImagePart imagePart = (ImagePart)temp.ContentPart.GetPartById(temp.ResourceID);
                        newContentPart.AddPart<ImagePart>(imagePart);
                        imageReference.Attribute(attributeName).Value = newContentPart.GetIdOfPart(imagePart);
                    }
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                try
                {
                    ExternalRelationship er = oldContentPart.GetExternalRelationship(relId);
                    ExternalRelationship newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    imageReference.Attribute(R.id).Value = newEr.Id;
                }
                catch (KeyNotFoundException)
                {
                    throw new PresentationBuilderInternalException("Source {0} is unsupported document - contains reference to NULL image");
                }
            }
        }

        // General function for handling images that tries to use an existing image if they are the same
        private static ImageData ManageImageCopy(ImagePart oldImage, OpenXmlPart newContentPart, List<ImageData> images)
        {
            ImageData oldImageData = new ImageData(newContentPart, oldImage);
            foreach (ImageData item in images)
            {
//                if (newContentPart != item.ContentPart)
//                    continue;
                if (item.Compare(oldImageData))
                    return item;
            }
            images.Add(oldImageData);
            return oldImageData;
        }

        private static void CopyRelatedSound(PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart,
            XElement soundReference, XName attributeName)
        {
            string relId = (string)soundReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;
            AudioReferenceRelationship temp = (AudioReferenceRelationship)oldContentPart.GetReferenceRelationship(relId);
            MediaDataPart newSound = newDocument.CreateMediaDataPart(temp.DataPart.ContentType);
            newSound.FeedData(temp.DataPart.GetStream());
            AudioReferenceRelationship newRel = null;
            if (newContentPart is SlidePart)
                newRel = ((SlidePart)newContentPart).AddAudioReferenceRelationship(newSound);
            if (newContentPart is SlideLayoutPart)
                newRel = ((SlideLayoutPart)newContentPart).AddAudioReferenceRelationship(newSound);
            if (newContentPart is SlideMasterPart)
                newRel = ((SlideMasterPart)newContentPart).AddAudioReferenceRelationship(newSound);
            soundReference.Attribute(attributeName).Value = newRel.Id;
        }
    }

    public class PresentationBuilderException : Exception
    {
        public PresentationBuilderException(string message) : base(message) { }
    }

    public class PresentationBuilderInternalException : Exception
    {
        public PresentationBuilderInternalException(string message) : base(message) { }
    }
}
