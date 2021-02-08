using System;
using System.Net;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var file = "test.docx";
using var doc = WordprocessingDocument.Create(file, WordprocessingDocumentType.Document);
var mainDocumentPart = doc.AddMainDocumentPart();

mainDocumentPart.Document = new Document();

mainDocumentPart.Document.Body = new Body();

var body = mainDocumentPart.Document.Body;
string[] contents = {"left", "centre", "right"};


var pageWidth = 12240;
var marginLeft = 1440;
var marginRight = 1440;
var p = new Paragraph();
var paragraphProperties = new ParagraphProperties(
    new Tabs(
        new TabStop(){Val = TabStopValues.Left, Position = 0},
        new TabStop(){Val = TabStopValues.Center, Position = (Int32Value)((pageWidth - marginLeft - marginRight)/2)},
        new TabStop(){Val = TabStopValues.Right, Position = pageWidth - marginLeft - marginRight}
    ));
p.AppendChild(paragraphProperties.CloneNode(true));

p.AppendChild(new Run(new Text(contents[0])));
p.AppendChild(new Run(new TabChar()));
p.AppendChild(new Run(new Text(contents[1])));
p.AppendChild(new Run(new TabChar()));
p.AppendChild(new Run(new Text(contents[2])));

body.AppendChild(p);


// case 2 the text is too longggg
string[] line2 = {"left2", "centre 2 long text that overlap the right block, why you are so long, long long  longggggggggggggggggggggggggggggggggggggg", "right2"};
var p2 = new Paragraph();
p2.AppendChild(paragraphProperties.CloneNode(true));

p2.AppendChild(new Run(new Text(line2[0])));
p2.AppendChild(new Run(new TabChar()));
p2.AppendChild(new Run(new Text(line2[1])));
p2.AppendChild(new Run(new TabChar()));
p2.AppendChild(new Run(new Text(line2[2])));

body.AppendChild(p2);

body.AppendChild(new SectionProperties(
    new PageSize(){Width = 12240, Height = 15840},
    new PageMargin(){Top = 1440, Bottom = 1440, Right = 1440, Left = 1440, Header = 720, Footer = 720},
    new Columns(){Space = "720"},
    new DocGrid(){LinePitch = 360}));

doc.Save();  //see test.docx


