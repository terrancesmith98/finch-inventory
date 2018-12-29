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
    }
}