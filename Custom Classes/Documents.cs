using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.DocumentObjectModel.Shapes;
using Finch_Inventory.Models;
using System.IO;

namespace Finch_Inventory.Custom_Classes
{
    public class Documents
    {
        internal static Document WeeklyInventoryAudit(List<Clothing> clothings)
        {
            //create new Migradoc document
            Document document = new Document();
            document.Info.Title = "Clothing/Roll Inventory Audit Report";
            document.Info.Subject = "Displays a inventory per Paper Machine.";
            document.Info.Author = "Terry Smith Custom Applications";
            Styles.DefineStyles(document);
            DefineWeeklyContentSection(document);

            //add report heading
            document.LastSection.AddParagraph("Clothing Inventory as of  - " + DateTime.Now.ToShortDateString(), "Heading1");
            document.LastSection.AddParagraph("", "FooterText");

            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Machine 1", "FooterText");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings1 = clothings.Where(x => x.PM_Number == 1).ToList();
            Table table1 = Tables.BuildInventoryAuditTable(clothings1);
            document.LastSection.Add(table1);

            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Machine 2", "FooterText");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings2 = clothings.Where(x => x.PM_Number == 2).ToList();
            Table table2 = Tables.BuildInventoryAuditTable(clothings2);
            document.LastSection.Add(table2);

            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Machine 3", "FooterText");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings3 = clothings.Where(x => x.PM_Number == 3).ToList();
            Table table3 = Tables.BuildInventoryAuditTable(clothings3);
            document.LastSection.Add(table3);

            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Machine 4", "FooterText");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings4 = clothings.Where(x => x.PM_Number == 4).ToList();
            Table table4 = Tables.BuildInventoryAuditTable(clothings4);
            document.LastSection.Add(table4);

            return document;
        }

        private static void DefineWeeklyContentSection(Document document)
        {
            Section section = document.AddSection();
            section.PageSetup.HeaderDistance = 2;
            section.PageSetup.LeftMargin = 10;
            section.PageSetup.TopMargin = 30;
            section.PageSetup.DifferentFirstPageHeaderFooter = true;


            //Create report footer
            HeaderFooter footer = section.Footers.FirstPage;
            Paragraph paragraph = new Paragraph();
            paragraph.AddFormattedText("Report Printed on: " + DateTime.Now.ToShortDateString(), "FooterText");
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            footer.Add(paragraph);


        }
    }
}