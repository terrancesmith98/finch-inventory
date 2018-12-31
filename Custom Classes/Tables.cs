using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Linq;
using System.Collections.Generic;
using Finch_Inventory.Models;

namespace Finch_Inventory.Custom_Classes
{
    public class Tables
    {
        static FinchDbContext db = new FinchDbContext();

        internal static Table BuildInventoryAuditTable(List<Clothing> clothings)
        {
            Table auditTable = new Table();
            auditTable.Style = "Table";
            auditTable.Borders.Width = .075;
            auditTable.Rows.Alignment = RowAlignment.Center;

            //define header columns
            Column posCol = auditTable.AddColumn(Unit.FromCentimeter(6));
            posCol.Format.Alignment = ParagraphAlignment.Left;
            posCol.Format.Font.Size = 6.5;

            Column typeCol = auditTable.AddColumn(Unit.FromCentimeter(8));
            typeCol.Format.Alignment = ParagraphAlignment.Center;
            typeCol.Format.Font.Size = 6.5;

            Column locCol = auditTable.AddColumn(Unit.FromCentimeter(3));
            locCol.Format.Alignment = ParagraphAlignment.Center;
            locCol.Format.Font.Size = 6.5;

            //define headers row
            Row headers = auditTable.AddRow();
            headers.Shading.Color = Colors.LightGray;
            headers.Height = 10;

            //define header row cells
            Cell posHeader = headers.Cells[0];
            posHeader.Format.Font.Size = 6.5;
            posHeader.VerticalAlignment = VerticalAlignment.Center;
            posHeader.AddParagraph("Position");

            Cell typeHeader = headers.Cells[1];
            typeHeader.Format.Font.Size = 6.5;
            typeHeader.VerticalAlignment = VerticalAlignment.Center;
            typeHeader.AddParagraph("Type");

            Cell locHeader = headers.Cells[2];
            locHeader.Format.Font.Size = 6.5;
            locHeader.VerticalAlignment = VerticalAlignment.Center;
            locHeader.AddParagraph("Location");

            //fill content cells
            var rowCounter = 1;
            foreach (var clothing in clothings)
            {
                Row r = auditTable.AddRow();
                r.Height = 13;
                if (rowCounter % 2 == 0)
                {
                    r.Format.Shading.Color = Colors.LightGray;
                }
                

                Cell cPos = r.Cells[0];
                cPos.Format.Font.Size = 7;
                cPos.AddParagraph($"{clothing.Position.Position1} - {clothing.Dimensions}");
                cPos.VerticalAlignment = VerticalAlignment.Center;

                Cell cType = r.Cells[1];
                cType.Format.Font.Size = 7;
                cType.AddParagraph($"{clothing.Manufacturer} - #{clothing.Serial_Number} - {clothing.Type.Type1}");
                cType.VerticalAlignment = VerticalAlignment.Center;

                Cell cLoc = r.Cells[2];
                cLoc.Format.Font.Size = 7;
                cLoc.AddParagraph(clothing.Location.Location1);
                cLoc.VerticalAlignment = VerticalAlignment.Center;
                rowCounter++;
            }

            return auditTable;
        }

        internal static Table BuildWeeklyPMTable(List<Clothing> clothings)
        {
            Table weeklyPMTable = new Table();
            weeklyPMTable.Style = "Table";
            weeklyPMTable.Borders.Width = .075;
            weeklyPMTable.Rows.Alignment = RowAlignment.Center;

            //define header columns
            Column posCol = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            posCol.Format.Alignment = ParagraphAlignment.Left;
            posCol.Format.Font.Size = 6.5;

            Column pm1Col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Col.Format.Alignment = ParagraphAlignment.Center;
            pm1Col.Format.Font.Size = 6.5;

            Column pm1Colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colb.Format.Alignment = ParagraphAlignment.Center;
            pm1Colb.Format.Font.Size = 6.5;

            Column pm1Colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colc.Format.Alignment = ParagraphAlignment.Center;
            pm1Colc.Format.Font.Size = 6.5;

            Column pm2col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2col.Format.Alignment = ParagraphAlignment.Center;
            pm2col.Format.Font.Size = 6.5;

            Column pm2colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colb.Format.Alignment = ParagraphAlignment.Center;
            pm2colb.Format.Font.Size = 6.5;

            Column pm2colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colc.Format.Alignment = ParagraphAlignment.Center;
            pm2colc.Format.Font.Size = 6.5;

            Column pm3col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3col.Format.Alignment = ParagraphAlignment.Center;
            pm3col.Format.Font.Size = 6.5;

            Column pm3colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colb.Format.Alignment = ParagraphAlignment.Center;
            pm3colb.Format.Font.Size = 6.5;

            Column pm3colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colc.Format.Alignment = ParagraphAlignment.Center;
            pm3colc.Format.Font.Size = 6.5;

            Column pm4col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4col.Format.Alignment = ParagraphAlignment.Center;
            pm4col.Format.Font.Size = 6.5;

            Column pm4colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colb.Format.Alignment = ParagraphAlignment.Center;
            pm4colb.Format.Font.Size = 6.5;

            Column pm4colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colc.Format.Alignment = ParagraphAlignment.Center;
            pm4colc.Format.Font.Size = 6.5;

            //define headers row
            Row headers = weeklyPMTable.AddRow();
            headers.Shading.Color = Colors.LightGray;
            headers.Height = 15;

            //define header row cells
            Cell posHeader = headers.Cells[0];
            posHeader.Format.Font.Size = 6.5;
            posHeader.VerticalAlignment = VerticalAlignment.Center;
            posHeader.MergeDown = 2;
            posHeader.AddParagraph("Position");

            Cell pm1Header = headers.Cells[1];
            pm1Header.Format.Font.Size = 6.5;
            pm1Header.MergeRight = 3;
            pm1Header.VerticalAlignment = VerticalAlignment.Center;
            pm1Header.AddParagraph("#1 PM");

            Cell pm2Header = headers.Cells[2];
            pm2Header.Format.Font.Size = 6.5;
            pm2Header.MergeRight = 3;
            pm2Header.VerticalAlignment = VerticalAlignment.Center;
            pm2Header.AddParagraph("#2 PM");

            Cell pm3Header = headers.Cells[3];
            pm3Header.Format.Font.Size = 6.5;
            pm3Header.MergeRight = 3;
            pm3Header.VerticalAlignment = VerticalAlignment.Center;
            pm3Header.AddParagraph("#3 PM");

            Cell pm4Header = headers.Cells[1];
            pm4Header.Format.Font.Size = 6.5;
            pm4Header.MergeRight = 3;
            pm4Header.VerticalAlignment = VerticalAlignment.Center;
            pm4Header.AddParagraph("#4 PM");

            //define header row 2
            Row headers2 = weeklyPMTable.AddRow();
            headers2.Height = 10;

            //define 2nd header row cells
            //Cell pos2Header = headers2.Cells[0];
            //pos2Header.Format.Font.Size = 6.5;
            //pos2Header.VerticalAlignment = VerticalAlignment.Center;
            //pos2Header.AddParagraph("Position");


            return weeklyPMTable;
        }
    }
}