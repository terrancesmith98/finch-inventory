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
            Column posCol = auditTable.AddColumn(Unit.FromCentimeter(3));
            posCol.Format.Alignment = ParagraphAlignment.Left;
            posCol.Format.Font.Size = 6.5;

            Column macCol = auditTable.AddColumn(Unit.FromCentimeter(3));
            macCol.Format.Alignment = ParagraphAlignment.Center;
            macCol.Format.Font.Size = 6.5;

            Column typeCol = auditTable.AddColumn(Unit.FromCentimeter(3));
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

            Cell macHeader = headers.Cells[0];
            macHeader.Format.Font.Size = 6.5;
            macHeader.VerticalAlignment = VerticalAlignment.Center;
            macHeader.AddParagraph("Position");

            Cell typeHeader = headers.Cells[0];
            typeHeader.Format.Font.Size = 6.5;
            typeHeader.VerticalAlignment = VerticalAlignment.Center;
            typeHeader.AddParagraph("Position");

            Cell locHeader = headers.Cells[0];
            locHeader.Format.Font.Size = 6.5;
            locHeader.VerticalAlignment = VerticalAlignment.Center;
            locHeader.AddParagraph("Position");

            //fill content cells
            foreach (var clothing in clothings)
            {
                Row r = auditTable.AddRow();
                r.Height = 13;

                Cell cPos = r.Cells[0];
                cPos.Format.Font.Size = 7;
                cPos.AddParagraph(clothing.Position.Position1);
                cPos.VerticalAlignment = VerticalAlignment.Center;

                Cell cMac = r.Cells[1];
                cMac.Format.Font.Size = 7;
                cMac.AddParagraph(clothing.Machine.Machine1);
                cMac.VerticalAlignment = VerticalAlignment.Center;

                Cell cType = r.Cells[2];
                cType.Format.Font.Size = 7;
                cType.AddParagraph(clothing.Type.Type1);
                cType.VerticalAlignment = VerticalAlignment.Center;
            }

            return auditTable;
        }
    }
}