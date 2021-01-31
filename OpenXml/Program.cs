using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml;
using Extensions = OpenXml.Extensions;

var ls = new SectionProperties(
    new PageSize()
    {
        Orient = PageOrientationValues.Landscape,
        Height = 12240,
        Width = 15840

    }
);

var pageProp = new ParagraphProperties()
{
    SectionProperties = ls
};

var pt = new SectionProperties(
    new PageSize()
    {
        Orient = PageOrientationValues.Portrait,
        Width = 12240,
        Height = 15840

    }
);

var pagePropPt = new ParagraphProperties()
{
    SectionProperties = pt
};




string path = $"./test-{DateTime.Now.ToString("hhmmss")}.docx";
string txt = "../../../test.txt";

var blockContents = File.ReadAllLines(txt); // read data from database

// now we create a file from lines in database
using var wordprocessingDocument = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
MainDocumentPart mainDocument = wordprocessingDocument.AddMainDocumentPart();
wordprocessingDocument.MainDocumentPart.Document = new Document();
var documentBody = new Body();
mainDocument.Document.AddChild(documentBody);

// each line is a list of paragraph (blockContent from Database)
int index = 1;
documentBody.Append(new Body(@$"<w:body xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">{string.Join(' ', blockContents)}</w:body>").Elements().Select(e =>
{
    e.SetAttribute(Extensions.BioContentId, $"{index++}");
    return e.CloneNode(true);
}));
// wordprocessingDocument.Save();
// wordprocessingDocument.Close();
// using var editDocument = WordprocessingDocument.Open(path, true);
// var documentBody = editDocument.MainDocumentPart.Document.Body;

// TODO: find element with bioContentId = 5 or 7
var bioContentIds = new[]{5, 7};

/*
 * The way to inserting page break
 * 1. From beginning to first TFL => portrait
 * 2. First TFL => landscape
 * 3. Each TFL => 1 page
 *
 * To adapt the requirement
 * 1. 
 * - When see the 1st TFL, we set the previous <p> is portrait page => this will set all previous page to portrait
 * - Insert last render page break to current <p> => this will stop the setting go to next section
 * 2.
 * - Add the default page format is landscape, it will apply to all other pages
 * -- Insert/Edit last node of <body> to be a landscape format
 * 3. 
 * - When see second TFL and from that point to last TFL, insert last render page break and <br type=page>
 * - The first TFL is skipped this <br type=page>
 */

bool firstTLF = false;
foreach (var paragraph in documentBody.ChildElements)
{
    if (bioContentIds.Contains(paragraph.GetBioAttribute())) // begin of a TFL
    {
        if (!firstTLF)
        {
            // do with session break;
            var previousSibling = paragraph.PreviousSibling();
            var oldSessionProperties = previousSibling.ChildElements.FirstOrDefault(e => e is ParagraphProperties);
            if (oldSessionProperties == null)
            {
                previousSibling.InsertAt(pagePropPt.CloneNode(true), 0);
            }
            else
            {
                previousSibling.ReplaceChild(pagePropPt.CloneNode(true), oldSessionProperties); // add session break =)
            }
            
            paragraph.FirstChild.InsertAt(new LastRenderedPageBreak(), 0);

        }
        else
        {
            //do with page break;
            paragraph.FirstChild.InsertAt(new LastRenderedPageBreak(), 0);
            paragraph.FirstChild.InsertAt(new Break()
            {
                Type = BreakValues.Page
            }, 1);
        }

        firstTLF = true;
        
    }
}

if (documentBody.LastChild is SectionProperties)
{
    documentBody.RemoveChild(documentBody.LastChild);
}
documentBody.AppendChild(ls.CloneNode(true)); // add Landscape as default

wordprocessingDocument.Save(); // see file in bin/debug/net.5.0/test-xxxx.docx