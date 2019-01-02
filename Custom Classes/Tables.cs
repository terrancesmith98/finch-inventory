using Finch_Inventory.Models;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Linq;

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
            var weeklyPMTable = new Table
            {
                Style = "Table"
            };
            weeklyPMTable.Borders.Width = .075;
            weeklyPMTable.Rows.Alignment = RowAlignment.Center;

            #region Header Columns
            //define header columns
            var posCol = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            posCol.Format.Alignment = ParagraphAlignment.Center;
            posCol.Format.Font.Size = 6.5;

            var pm1Col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Col.Format.Alignment = ParagraphAlignment.Center;
            pm1Col.Format.Font.Size = 6.5;

            var pm1Colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colb.Format.Alignment = ParagraphAlignment.Center;
            pm1Colb.Format.Font.Size = 6.5;

            var pm1Colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colc.Format.Alignment = ParagraphAlignment.Center;
            pm1Colc.Format.Font.Size = 6.5;

            var pm2col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2col.Format.Alignment = ParagraphAlignment.Center;
            pm2col.Format.Font.Size = 6.5;

            var pm2colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colb.Format.Alignment = ParagraphAlignment.Center;
            pm2colb.Format.Font.Size = 6.5;

            var pm2colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colc.Format.Alignment = ParagraphAlignment.Center;
            pm2colc.Format.Font.Size = 6.5;

            var pm3col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3col.Format.Alignment = ParagraphAlignment.Center;
            pm3col.Format.Font.Size = 6.5;

            var pm3colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colb.Format.Alignment = ParagraphAlignment.Center;
            pm3colb.Format.Font.Size = 6.5;

            var pm3colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colc.Format.Alignment = ParagraphAlignment.Center;
            pm3colc.Format.Font.Size = 6.5;

            var pm4col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4col.Format.Alignment = ParagraphAlignment.Center;
            pm4col.Format.Font.Size = 6.5;

            var pm4colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colb.Format.Alignment = ParagraphAlignment.Center;
            pm4colb.Format.Font.Size = 6.5;

            var pm4colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colc.Format.Alignment = ParagraphAlignment.Center;
            pm4colc.Format.Font.Size = 6.5;

            #endregion

            #region Header Rows

            //define headers row
            var headers = weeklyPMTable.AddRow();
            headers.Shading.Color = Colors.LightGray;
            headers.Height = 15;

            //define header row cells
            var posHeader = headers.Cells[0];
            posHeader.Format.Font.Size = 6.5;
            posHeader.VerticalAlignment = VerticalAlignment.Center;
            posHeader.MergeDown = 1;
            posHeader.AddParagraph("Position");

            var pm1Header = headers.Cells[1];
            pm1Header.Format.Font.Size = 6.5;
            pm1Header.MergeRight = 2;
            pm1Header.VerticalAlignment = VerticalAlignment.Center;
            pm1Header.AddParagraph("#1 PM");

            var pm2Header = headers.Cells[4];
            pm2Header.Format.Font.Size = 6.5;
            pm2Header.MergeRight = 2;
            pm2Header.VerticalAlignment = VerticalAlignment.Center;
            pm2Header.AddParagraph("#2 PM");

            var pm3Header = headers.Cells[7];
            pm3Header.Format.Font.Size = 6.5;
            pm3Header.MergeRight = 2;
            pm3Header.VerticalAlignment = VerticalAlignment.Center;
            pm3Header.AddParagraph("#3 PM");

            var pm4Header = headers.Cells[10];
            pm4Header.Format.Font.Size = 6.5;
            pm4Header.MergeRight = 2;
            pm4Header.VerticalAlignment = VerticalAlignment.Center;
            pm4Header.AddParagraph("#4 PM");

            //define header row 2
            var headers2 = weeklyPMTable.AddRow();
            headers2.Height = 10;

            //define 2nd header row cells
            var pm1PastHeader = headers2.Cells[1];
            pm1PastHeader.Format.Font.Size = 6.5;
            pm1PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1PastHeader.AddParagraph("PAST");

            var pm1CurrentHeader = headers2.Cells[2];
            pm1CurrentHeader.Format.Font.Size = 6.5;
            pm1CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1CurrentHeader.AddParagraph("CURRENT");

            var pm1GoalHeader = headers2.Cells[3];
            pm1GoalHeader.Format.Font.Size = 6.5;
            pm1GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1GoalHeader.AddParagraph("GOAL");

            var pm2PastHeader = headers2.Cells[4];
            pm2PastHeader.Format.Font.Size = 6.5;
            pm2PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2PastHeader.AddParagraph("PAST");

            var pm2CurrentHeader = headers2.Cells[5];
            pm2CurrentHeader.Format.Font.Size = 6.5;
            pm2CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2CurrentHeader.AddParagraph("CURRENT");

            var pm2GoalHeader = headers2.Cells[6];
            pm2GoalHeader.Format.Font.Size = 6.5;
            pm2GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2GoalHeader.AddParagraph("GOAL");

            var pm3PastHeader = headers2.Cells[7];
            pm3PastHeader.Format.Font.Size = 6.5;
            pm3PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3PastHeader.AddParagraph("PAST");

            var pm3CurrentHeader = headers2.Cells[8];
            pm3CurrentHeader.Format.Font.Size = 6.5;
            pm3CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3CurrentHeader.AddParagraph("CURRENT");

            var pm3GoalHeader = headers2.Cells[9];
            pm3GoalHeader.Format.Font.Size = 6.5;
            pm3GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3GoalHeader.AddParagraph("GOAL");

            var pm4PastHeader = headers2.Cells[10];
            pm4PastHeader.Format.Font.Size = 6.5;
            pm4PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4PastHeader.AddParagraph("PAST");

            var pm4CurrentHeader = headers2.Cells[11];
            pm4CurrentHeader.Format.Font.Size = 6.5;
            pm4CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4CurrentHeader.AddParagraph("CURRENT");

            var pm4GoalHeader = headers2.Cells[12];
            pm4GoalHeader.Format.Font.Size = 6.5;
            pm4GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4GoalHeader.AddParagraph("GOAL");
            #endregion

            //define Content Rows
            #region Wire Position
            // ----- Wire Position Row ----- //
            var clothingWire1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 1);
            var clothingWire2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 2);
            var clothingWire3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 3);
            var clothingWire4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 4);
            var rowWire = weeklyPMTable.AddRow();
            rowWire.Height = 10;
            var rowWire2 = weeklyPMTable.AddRow();
            rowWire2.Height = 10;


            //define wire position row cells

            var cellWire = rowWire.Cells[0];
            cellWire.Format.Font.Size = 6.5;
            cellWire.VerticalAlignment = VerticalAlignment.Center;
            cellWire.MergeDown = 1;
            cellWire.AddParagraph("WIRE");

            #region Wire PM 1
            // ----- Wire Position PM 1 ----- //
            var cellWire1Past = rowWire.Cells[1];
            cellWire1Past.Format.Font.Size = 6.5;
            cellWire1Past.VerticalAlignment = VerticalAlignment.Center;
            cellWire1Past.AddParagraph($"100");

            var cellWire1Current = rowWire.Cells[2];
            cellWire1Current.Format.Font.Size = 6.5;
            cellWire1Current.VerticalAlignment = VerticalAlignment.Center;
            var cellWire1Age = 0;
            if (clothingWire1 != null)
            {
                cellWire1Age = clothingWire1.Age != null ? Convert.ToInt32(clothingWire1.Age) : 0;
            }
            
            cellWire1Current.AddParagraph($"{cellWire1Age}");

            var cellWire1Goal = rowWire.Cells[3];
            cellWire1Goal.Format.Font.Size = 6.5;
            cellWire1Goal.VerticalAlignment = VerticalAlignment.Center;
            var wire1Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 2);
            cellWire1Goal.AddParagraph($"{wire1Goal.Goal1}");

            var cellWire1Info = rowWire2.Cells[1];
            cellWire1Info.Format.Font.Size = 6.5;
            cellWire1Info.MergeRight = 2;
            cellWire1Info.Shading.Color = Colors.LightBlue;
            if (clothingWire1 != null)
            {
                cellWire1Info.AddParagraph($"{clothingWire1.Manufacturer} {clothingWire1.Serial_Number}");
            }
            else cellWire1Info.AddParagraph("NA");
            #endregion

            #region Wire PM 2

            // ----- Wire Position PM 2 ----- //
            var cellWire2Past = rowWire.Cells[4];
            cellWire2Past.Format.Font.Size = 6.5;
            cellWire2Past.VerticalAlignment = VerticalAlignment.Center;
            cellWire2Past.AddParagraph($"100");

            var cellWire2Current = rowWire.Cells[5];
            cellWire2Current.Format.Font.Size = 6.5;
            cellWire2Current.VerticalAlignment = VerticalAlignment.Center;
            var cellWire2Age = 0;
            if (clothingWire2 != null)
            {
                cellWire2Age = clothingWire2.Age != null ? Convert.ToInt32(clothingWire2.Age) : 0;
            }

            cellWire2Current.AddParagraph($"{cellWire2Age}");

            var cellWire2Goal = rowWire.Cells[6];
            cellWire2Goal.Format.Font.Size = 6.5;
            cellWire2Goal.VerticalAlignment = VerticalAlignment.Center;
            var wire2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 2);
            cellWire2Goal.AddParagraph($"{wire2Goal.Goal1}");

            var cellWire2Info = rowWire2.Cells[4];
            cellWire2Info.Format.Font.Size = 6.5;
            cellWire2Info.MergeRight = 2;
            cellWire2Info.Shading.Color = Colors.LightBlue;
            if (clothingWire2 != null)
            {
                cellWire2Info.AddParagraph($"{clothingWire2.Manufacturer} {clothingWire2.Serial_Number}");
            }
            else cellWire2Info.AddParagraph("NA");
            #endregion

            #region Wire PM 3
            // ----- Wire Position PM 3 ----- //
            var cellWire3Past = rowWire.Cells[7];
            cellWire3Past.Format.Font.Size = 6.5;
            cellWire3Past.VerticalAlignment = VerticalAlignment.Center;
            cellWire3Past.AddParagraph($"100");

            var cellWire3Current = rowWire.Cells[8];
            cellWire3Current.Format.Font.Size = 6.5;
            cellWire3Current.VerticalAlignment = VerticalAlignment.Center;
            var cellWire3Age = 0;
            if (clothingWire3 != null)
            {
                cellWire3Age = clothingWire3.Age != null ? Convert.ToInt32(clothingWire3.Age) : 0;
            }

            cellWire3Current.AddParagraph($"{cellWire3Age}");

            var cellWire3Goal = rowWire.Cells[9];
            cellWire3Goal.Format.Font.Size = 6.5;
            cellWire3Goal.VerticalAlignment = VerticalAlignment.Center;
            var wire3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 2);
            cellWire3Goal.AddParagraph($"{wire3Goal.Goal1}");

            var cellWire3Info = rowWire2.Cells[7];
            cellWire3Info.Format.Font.Size = 6.5;
            cellWire3Info.MergeRight = 2;
            cellWire3Info.Shading.Color = Colors.LightBlue;
            if (clothingWire3 != null)
            {
                cellWire3Info.AddParagraph($"{clothingWire3.Manufacturer} {clothingWire3.Serial_Number}");
            }
            else cellWire3Info.AddParagraph("NA");
            #endregion

            #region Wire PM 4
            // ----- Wire Position PM 4 ----- //
            var cellWire4Past = rowWire.Cells[10];
            cellWire4Past.Format.Font.Size = 6.5;
            cellWire4Past.VerticalAlignment = VerticalAlignment.Center;
            cellWire4Past.AddParagraph($"100");

            var cellWire4Current = rowWire.Cells[11];
            cellWire4Current.Format.Font.Size = 6.5;
            cellWire4Current.VerticalAlignment = VerticalAlignment.Center;
            var cellWire4Age = 0;
            if (clothingWire4 != null)
            {
                cellWire4Age = clothingWire4.Age != null ? Convert.ToInt32(clothingWire4.Age) : 0;
            }

            cellWire4Current.AddParagraph($"{cellWire4Age}");

            var cellWire4Goal = rowWire.Cells[12];
            cellWire4Goal.Format.Font.Size = 6.5;
            cellWire4Goal.VerticalAlignment = VerticalAlignment.Center;
            var wire4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 2);
            cellWire4Goal.AddParagraph($"{wire4Goal.Goal1}");

            var cellWire4Info = rowWire2.Cells[10];
            cellWire4Info.Format.Font.Size = 6.5;
            cellWire4Info.MergeRight = 2;
            cellWire4Info.Shading.Color = Colors.LightBlue;
            if (clothingWire4 != null)
            {
                cellWire4Info.AddParagraph($"{clothingWire4.Manufacturer} {clothingWire4.Serial_Number}");
            }
            else cellWire4Info.AddParagraph("NA");
            #endregion

            #endregion // Wire Position

            #region 1st Press Position

            // ----- 1st Press Position Row ----- //
            // add spacer row
            var spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            var cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingFirstPress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 1);
            var clothingFirstPress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 2);
            var clothingFirstPress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 3);
            var clothingFirstPress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 4);
            var rowFirstPress = weeklyPMTable.AddRow();
            rowFirstPress.Height = 10;
            var rowFirstPress2 = weeklyPMTable.AddRow();
            rowFirstPress2.Height = 10;

            //define 1st Press position row cells

            var cell1Press = rowFirstPress.Cells[0];
            cell1Press.Format.Font.Size = 6.5;
            cell1Press.VerticalAlignment = VerticalAlignment.Center;
            cell1Press.MergeDown = 1;
            cell1Press.AddParagraph("1ST PRESS");

            #region 1st Press PM 1
            // ----- 1st Press Position PM 1 ----- //
            var cell1Press1Past = rowFirstPress.Cells[1];
            cell1Press1Past.Format.Font.Size = 6.5;
            cell1Press1Past.VerticalAlignment = VerticalAlignment.Center;
            cell1Press1Past.AddParagraph($"100");

            var cell1Press1Current = rowFirstPress.Cells[2];
            cell1Press1Current.Format.Font.Size = 6.5;
            cell1Press1Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPress1Age = 0;
            if (clothingFirstPress1 != null)
            {
                firstPress1Age = clothingFirstPress1.Age != null ? Convert.ToInt32(clothingFirstPress1.Age) : 0;
            }

            cell1Press1Current.AddParagraph($"{firstPress1Age}");

            var cell1Press1Goal = rowFirstPress.Cells[3];
            cell1Press1Goal.Format.Font.Size = 6.5;
            cell1Press1Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPress1Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 3);
            cell1Press1Goal.AddParagraph($"{firstPress1Goal.Goal1}");

            var cell1Press1Info = rowFirstPress2.Cells[1];
            cell1Press1Info.Format.Font.Size = 6.5;
            cell1Press1Info.MergeRight = 2;
            cell1Press1Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstPress1 != null)
            {
                cell1Press1Info.AddParagraph($"{clothingFirstPress1.Manufacturer} {clothingFirstPress1.Serial_Number}");
            }
            else cell1Press1Info.AddParagraph("NA");
            #endregion

            #region 1st Press PM 2
            // ----- 1st Press Position PM 2 ----- //
            var cell1Press2Past = rowFirstPress.Cells[4];
            cell1Press2Past.Format.Font.Size = 6.5;
            cell1Press2Past.VerticalAlignment = VerticalAlignment.Center;
            cell1Press2Past.AddParagraph($"100");

            var cell1Press2Current = rowFirstPress.Cells[5];
            cell1Press2Current.Format.Font.Size = 6.5;
            cell1Press2Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPress2Age = 0;
            if (clothingFirstPress2 != null)
            {
                firstPress2Age = clothingFirstPress2.Age != null ? Convert.ToInt32(clothingFirstPress2.Age) : 0;
            }

            cell1Press2Current.AddParagraph($"{firstPress2Age}");

            var cell1Press2Goal = rowFirstPress.Cells[6];
            cell1Press2Goal.Format.Font.Size = 6.5;
            cell1Press2Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPress2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 3);
            cell1Press2Goal.AddParagraph($"{firstPress2Goal.Goal1}");

            var cell1Press2Info = rowFirstPress2.Cells[4];
            cell1Press2Info.Format.Font.Size = 6.5;
            cell1Press2Info.MergeRight = 2;
            cell1Press2Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstPress2 != null)
            {
                cell1Press2Info.AddParagraph($"{clothingFirstPress2.Manufacturer} {clothingFirstPress2.Serial_Number}");
            }
            else cell1Press2Info.AddParagraph("NA");
            #endregion

            #region 1st Press PM 3
            // ----- 1st Press Position PM 3 ----- //
            var cell1Press3Past = rowFirstPress.Cells[7];
            cell1Press3Past.Format.Font.Size = 6.5;
            cell1Press3Past.VerticalAlignment = VerticalAlignment.Center;
            cell1Press3Past.AddParagraph($"100");

            var cell1Press3Current = rowFirstPress.Cells[8];
            cell1Press3Current.Format.Font.Size = 6.5;
            cell1Press3Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPress3Age = 0;
            if (clothingFirstPress3 != null)
            {
                firstPress3Age = clothingFirstPress3.Age != null ? Convert.ToInt32(clothingFirstPress3.Age) : 0;
            }

            cell1Press3Current.AddParagraph($"{firstPress3Age}");

            var cell1Press3Goal = rowFirstPress.Cells[9];
            cell1Press3Goal.Format.Font.Size = 6.5;
            cell1Press3Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPress3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 3);
            cell1Press3Goal.AddParagraph($"{firstPress3Goal.Goal1}");

            var cell1Press3Info = rowFirstPress2.Cells[7];
            cell1Press3Info.Format.Font.Size = 6.5;
            cell1Press3Info.MergeRight = 2;
            cell1Press3Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstPress3 != null)
            {
                cell1Press3Info.AddParagraph($"{clothingFirstPress3.Manufacturer} {clothingFirstPress3.Serial_Number}");
            }
            else cell1Press3Info.AddParagraph("NA");
            #endregion

            #region 1st Press PM 4
            // ----- 1st Press Position PM 4 ----- //
            var cell1Press4Past = rowFirstPress.Cells[10];
            cell1Press4Past.Format.Font.Size = 6.5;
            cell1Press4Past.VerticalAlignment = VerticalAlignment.Center;
            cell1Press4Past.AddParagraph($"100");

            var cell1Press4Current = rowFirstPress.Cells[11];
            cell1Press4Current.Format.Font.Size = 6.5;
            cell1Press4Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPress4Age = 0;
            if (clothingFirstPress4 != null)
            {
                firstPress4Age = clothingFirstPress4.Age != null ? Convert.ToInt32(clothingFirstPress4.Age) : 0;
            }

            cell1Press4Current.AddParagraph($"{firstPress4Age}");

            var cell1Press4Goal = rowFirstPress.Cells[12];
            cell1Press4Goal.Format.Font.Size = 6.5;
            cell1Press4Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPress4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 3);
            cell1Press4Goal.AddParagraph($"{firstPress4Goal.Goal1}");

            var cell1Press4Info = rowFirstPress2.Cells[10];
            cell1Press4Info.Format.Font.Size = 6.5;
            cell1Press4Info.MergeRight = 2;
            cell1Press4Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstPress4 != null)
            {
                cell1Press4Info.AddParagraph($"{clothingFirstPress4.Manufacturer} {clothingFirstPress4.Serial_Number}");
            }
            else cell1Press4Info.AddParagraph("NA");
            #endregion

            #endregion //1st Press Position

            #region 2nd Press Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSecondPress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 1);
            var clothingSecondPress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 2);
            var clothingSecondPress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 3);
            var clothingSecondPress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 4);
            var rowSecondPress = weeklyPMTable.AddRow();
            rowSecondPress.Height = 10;
            var rowSecondPress2 = weeklyPMTable.AddRow();
            rowSecondPress2.Height = 10;

            //define 2nd Press position row cells

            var cell2Press = rowSecondPress.Cells[0];
            cell2Press.Format.Font.Size = 6.5;
            cell2Press.VerticalAlignment = VerticalAlignment.Center;
            cell2Press.MergeDown = 1;
            cell2Press.AddParagraph("2ND PRESS");

            #region 2nd Press PM 1
            // ----- 2nd Press Position PM 1 ----- //
            var cell2Press1Past = rowSecondPress.Cells[1];
            cell2Press1Past.Format.Font.Size = 6.5;
            cell2Press1Past.VerticalAlignment = VerticalAlignment.Center;
            cell2Press1Past.AddParagraph($"100");

            var cell2Press1Current = rowSecondPress.Cells[2];
            cell2Press1Current.Format.Font.Size = 6.5;
            cell2Press1Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPress1Age = 0;
            if (clothingSecondPress1 != null)
            {
                secondPress1Age = clothingSecondPress1.Age != null ? Convert.ToInt32(clothingSecondPress1.Age) : 0;
            }

            cell2Press1Current.AddParagraph($"{secondPress1Age}");

            var cell2Press1Goal = rowSecondPress.Cells[3];
            cell2Press1Goal.Format.Font.Size = 6.5;
            cell2Press1Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPress1Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 4);
            cell2Press1Goal.AddParagraph($"{secondPress1Goal.Goal1}");

            var cell2Press1Info = rowSecondPress2.Cells[1];
            cell2Press1Info.Format.Font.Size = 6.5;
            cell2Press1Info.MergeRight = 2;
            cell2Press1Info.Shading.Color = Colors.LightBlue;
            if (clothingSecondPress1 != null)
            {
                cell2Press1Info.AddParagraph($"{clothingSecondPress1.Manufacturer} {clothingSecondPress1.Serial_Number}");
            }
            else cell2Press1Info.AddParagraph("NA");
            #endregion

            #region 2nd Press PM 2
            // ----- 2nd Press Position PM 2 ----- //
            var cell2Press2Past = rowSecondPress.Cells[4];
            cell2Press2Past.Format.Font.Size = 6.5;
            cell2Press2Past.VerticalAlignment = VerticalAlignment.Center;
            cell2Press2Past.AddParagraph($"100");

            var cell2Press2Current = rowSecondPress.Cells[5];
            cell2Press2Current.Format.Font.Size = 6.5;
            cell2Press2Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPress2Age = 0;
            if (clothingSecondPress2 != null)
            {
                secondPress2Age = clothingSecondPress2.Age != null ? Convert.ToInt32(clothingSecondPress2.Age) : 0;
            }

            cell2Press2Current.AddParagraph($"{secondPress2Age}");

            var cell2Press2Goal = rowSecondPress.Cells[6];
            cell2Press2Goal.Format.Font.Size = 6.5;
            cell2Press2Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPress2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 4);
            cell2Press2Goal.AddParagraph($"{secondPress2Goal.Goal1}");

            var cell2Press2Info = rowSecondPress2.Cells[4];
            cell2Press2Info.Format.Font.Size = 6.5;
            cell2Press2Info.MergeRight = 2;
            cell2Press2Info.Shading.Color = Colors.LightBlue;
            if (clothingSecondPress2 != null)
            {
                cell2Press2Info.AddParagraph($"{clothingSecondPress2.Manufacturer} {clothingSecondPress2.Serial_Number}");
            }
            else cell2Press2Info.AddParagraph("NA");
            #endregion

            #region 2nd Press PM 3
            // ----- 2nd Press Position PM 3 ----- //
            var cell2Press3Past = rowSecondPress.Cells[7];
            cell2Press3Past.Format.Font.Size = 6.5;
            cell2Press3Past.VerticalAlignment = VerticalAlignment.Center;
            cell2Press3Past.AddParagraph($"100");

            var cell2Press3Current = rowSecondPress.Cells[8];
            cell2Press3Current.Format.Font.Size = 6.5;
            cell2Press3Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPress3Age = 0;
            if (clothingSecondPress3 != null)
            {
                secondPress3Age = clothingSecondPress3.Age != null ? Convert.ToInt32(clothingSecondPress3.Age) : 0;
            }

            cell2Press3Current.AddParagraph($"{secondPress3Age}");

            var cell2Press3Goal = rowSecondPress.Cells[9];
            cell2Press3Goal.Format.Font.Size = 6.5;
            cell2Press3Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPress3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 4);
            cell2Press3Goal.AddParagraph($"{secondPress3Goal.Goal1}");

            var cell2Press3Info = rowSecondPress2.Cells[7];
            cell2Press3Info.Format.Font.Size = 6.5;
            cell2Press3Info.MergeRight = 2;
            cell2Press3Info.Shading.Color = Colors.LightBlue;
            if (clothingSecondPress3 != null)
            {
                cell2Press3Info.AddParagraph($"{clothingSecondPress3.Manufacturer} {clothingSecondPress3.Serial_Number}");
            }
            else cell2Press3Info.AddParagraph("NA");
            #endregion

            #region 2nd Press PM 4
            // ----- 2nd Press Position PM 4 ----- //
            var cell2Press4Past = rowSecondPress.Cells[10];
            cell2Press4Past.Format.Font.Size = 6.5;
            cell2Press4Past.VerticalAlignment = VerticalAlignment.Center;
            cell2Press4Past.AddParagraph($"100");

            var cell2Press4Current = rowSecondPress.Cells[11];
            cell2Press4Current.Format.Font.Size = 6.5;
            cell2Press4Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPress4Age = 0;
            if (clothingSecondPress4 != null)
            {
                secondPress4Age = clothingSecondPress4.Age != null ? Convert.ToInt32(clothingSecondPress4.Age) : 0;
            }
            cell2Press4Current.AddParagraph($"{secondPress4Age}");

            var cell2Press4Goal = rowSecondPress.Cells[12];
            cell2Press4Goal.Format.Font.Size = 6.5;
            cell2Press4Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPress4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 4);
            cell2Press4Goal.AddParagraph($"{secondPress4Goal.Goal1}");

            var cell2Press4Info = rowSecondPress2.Cells[10];
            cell2Press4Info.Format.Font.Size = 6.5;
            cell2Press4Info.MergeRight = 2;
            cell2Press4Info.Shading.Color = Colors.LightBlue;
            if (clothingSecondPress4 != null)
            {
                cell2Press4Info.AddParagraph($"{clothingSecondPress4.Manufacturer} {clothingSecondPress4.Serial_Number}");
            }
            else cell2Press4Info.AddParagraph("NA");
            #endregion

            #endregion // 2nd Press Position

            #region 3rd Press Position

            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingThirdPress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 1);
            var clothingThirdPress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 2);
            var clothingThirdPress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 3);
            var clothingThirdPress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 4);
            var rowThirdPress = weeklyPMTable.AddRow();
            rowThirdPress.Height = 10;
            var rowThirdPress2 = weeklyPMTable.AddRow();
            rowThirdPress2.Height = 10;

            //define 3rd Press position row cells

            var cell3Press = rowThirdPress.Cells[0];
            cell3Press.Format.Font.Size = 6.5;
            cell3Press.VerticalAlignment = VerticalAlignment.Center;
            cell3Press.MergeDown = 1;
            cell3Press.AddParagraph("3rd PRESS");

            #region 3rd Press PM 1
            // ----- 3rd Press Position PM 1 ----- //
            var cell3Press1Past = rowThirdPress.Cells[1];
            cell3Press1Past.Format.Font.Size = 6.5;
            cell3Press1Past.VerticalAlignment = VerticalAlignment.Center;
            cell3Press1Past.AddParagraph($"100");

            var cell3Press1Current = rowThirdPress.Cells[2];
            cell3Press1Current.Format.Font.Size = 6.5;
            cell3Press1Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress1Age = 0;
            if (clothingThirdPress1 != null)
            {
                thirdPress1Age = clothingThirdPress1.Age != null ? Convert.ToInt32(clothingThirdPress1.Age) : 0;
            }

            cell3Press1Current.AddParagraph($"{thirdPress1Age}");

            var cell3Press1Goal = rowThirdPress.Cells[3];
            cell3Press1Goal.Format.Font.Size = 6.5;
            cell3Press1Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress1Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 5);
            cell3Press1Goal.AddParagraph($"{thirdPress1Goal.Goal1}");

            var cell3Press1Info = rowThirdPress2.Cells[1];
            cell3Press1Info.Format.Font.Size = 6.5;
            cell3Press1Info.MergeRight = 2;
            cell3Press1Info.Shading.Color = Colors.LightBlue;
            if (clothingThirdPress1 != null)
            {
                cell3Press1Info.AddParagraph($"{clothingThirdPress1.Manufacturer} {clothingThirdPress1.Serial_Number}");
            }
            else cell3Press1Info.AddParagraph("NA");
            #endregion

            #region 3rd Press PM 2
            // ----- 3rd Press Position PM 2 ----- //
            var cell3Press2Past = rowThirdPress.Cells[4];
            cell3Press2Past.Format.Font.Size = 6.5;
            cell3Press2Past.VerticalAlignment = VerticalAlignment.Center;
            cell3Press2Past.AddParagraph($"100");

            var cell3Press2Current = rowThirdPress.Cells[5];
            cell3Press2Current.Format.Font.Size = 6.5;
            cell3Press2Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress2Age = 0;
            if (clothingThirdPress2 != null)
            {
                thirdPress2Age = clothingThirdPress2.Age != null ? Convert.ToInt32(clothingThirdPress2.Age) : 0;
            }

            cell3Press2Current.AddParagraph($"{thirdPress2Age}");

            var cell3Press2Goal = rowThirdPress.Cells[6];
            cell3Press2Goal.Format.Font.Size = 6.5;
            cell3Press2Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 5);
            cell3Press2Goal.AddParagraph($"{thirdPress2Goal.Goal1}");

            var cell3Press2Info = rowThirdPress2.Cells[4];
            cell3Press2Info.Format.Font.Size = 6.5;
            cell3Press2Info.MergeRight = 2;
            cell3Press2Info.Shading.Color = Colors.LightBlue;
            if (clothingThirdPress2 != null)
            {
                cell3Press2Info.AddParagraph($"{clothingThirdPress2.Manufacturer} {clothingThirdPress2.Serial_Number}");
            }
            else cell3Press2Info.AddParagraph("NA");
            #endregion

            #region 3rd Press PM 3
            // ----- 3rd Press Position PM 3 ----- //
            var cell3Press3Past = rowThirdPress.Cells[7];
            cell3Press3Past.Format.Font.Size = 6.5;
            cell3Press3Past.VerticalAlignment = VerticalAlignment.Center;
            cell3Press3Past.AddParagraph($"100");

            var cell3Press3Current = rowThirdPress.Cells[8];
            cell3Press3Current.Format.Font.Size = 6.5;
            cell3Press3Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress3Age = 0;
            if (clothingThirdPress3 != null)
            {
                thirdPress3Age = clothingThirdPress3.Age != null ? Convert.ToInt32(clothingThirdPress3.Age) : 0;
            }

            cell3Press3Current.AddParagraph($"{thirdPress3Age}");

            var cell3Press3Goal = rowThirdPress.Cells[9];
            cell3Press3Goal.Format.Font.Size = 6.5;
            cell3Press3Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 5);
            cell3Press3Goal.AddParagraph($"{thirdPress3Goal.Goal1}");

            var cell3Press3Info = rowThirdPress2.Cells[7];
            cell3Press3Info.Format.Font.Size = 6.5;
            cell3Press3Info.MergeRight = 2;
            cell3Press3Info.Shading.Color = Colors.LightBlue;
            if (clothingThirdPress3 != null)
            {
                cell3Press3Info.AddParagraph($"{clothingThirdPress3.Manufacturer} {clothingThirdPress3.Serial_Number}");
            }
            else cell3Press3Info.AddParagraph("NA");
            #endregion

            #region 3rd Press PM 4
            // ----- 3rd Press Position PM 4 ----- //
            var cell3Press4Past = rowThirdPress.Cells[10];
            cell3Press4Past.Format.Font.Size = 6.5;
            cell3Press4Past.VerticalAlignment = VerticalAlignment.Center;
            cell3Press4Past.AddParagraph($"100");

            var cell3Press4Current = rowThirdPress.Cells[11];
            cell3Press4Current.Format.Font.Size = 6.5;
            cell3Press4Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress4Age = 0;
            if (clothingThirdPress4 != null)
            {
                thirdPress4Age = clothingThirdPress4.Age != null ? Convert.ToInt32(clothingThirdPress4.Age) : 0;
            }

            cell3Press4Current.AddParagraph($"{thirdPress4Age}");

            var cell3Press4Goal = rowThirdPress.Cells[12];
            cell3Press4Goal.Format.Font.Size = 6.5;
            cell3Press4Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPress4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 5);
            cell3Press4Goal.AddParagraph($"{thirdPress4Goal.Goal1}");

            var cell3Press4Info = rowThirdPress2.Cells[10];
            cell3Press4Info.Format.Font.Size = 6.5;
            cell3Press4Info.MergeRight = 2;
            cell3Press4Info.Shading.Color = Colors.LightBlue;
            if (clothingThirdPress4 != null)
            {
                cell3Press4Info.AddParagraph($"{clothingThirdPress4.Manufacturer} {clothingThirdPress4.Serial_Number}");
            }
            else cell3Press4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region 1st Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingFirstTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 1);
            var clothingFirstTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 2);
            var clothingFirstTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 3);
            var clothingFirstTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 4);
            var rowFirstTopDryer = weeklyPMTable.AddRow();
            rowFirstTopDryer.Height = 10;
            var rowFirstTopDryer2 = weeklyPMTable.AddRow();
            rowFirstTopDryer2.Height = 10;

            //define 1st Top position row cells

            var cellFirstTopDryer = rowFirstTopDryer.Cells[0];
            cellFirstTopDryer.Format.Font.Size = 6.5;
            cellFirstTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cellFirstTopDryer.MergeDown = 1;
            cellFirstTopDryer.AddParagraph("1ST TOP");

            #region 1st Top PM 1
            // ----- 1st Top Position PM 1 ----- //
            var cellFirstTopDryer1Past = rowFirstTopDryer.Cells[1];
            cellFirstTopDryer1Past.Format.Font.Size = 6.5;
            cellFirstTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cellFirstTopDryer1Past.AddParagraph($"100");

            var cellFirstTopDryer1Current = rowFirstTopDryer.Cells[2];
            cellFirstTopDryer1Current.Format.Font.Size = 6.5;
            cellFirstTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryer1Age = 0;
            if (clothingFirstTopDryer1 != null)
            {
                firstTopDryer1Age = clothingFirstTopDryer1.Age != null ? Convert.ToInt32(clothingFirstTopDryer1.Age) : 0;
            }

            cellFirstTopDryer1Current.AddParagraph($"{firstTopDryer1Age}");

            var cellFirstTopDryer1Goal = rowFirstTopDryer.Cells[3];
            cellFirstTopDryer1Goal.Format.Font.Size = 6.5;
            cellFirstTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 6);
            cellFirstTopDryer1Goal.AddParagraph($"{firstTopDryerlGoal.Goal1}");

            var cellFirstTopDryer1Info = rowFirstTopDryer2.Cells[1];
            cellFirstTopDryer1Info.Format.Font.Size = 6.5;
            cellFirstTopDryer1Info.MergeRight = 2;
            cellFirstTopDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstTopDryer1 != null)
            {
                cellFirstTopDryer1Info.AddParagraph($"{clothingFirstTopDryer1.Manufacturer} {clothingFirstTopDryer1.Serial_Number}");
            }
            else cellFirstTopDryer1Info.AddParagraph("NA");
            #endregion

            #region 1st Top PM 2
            // ----- 1st Top Position PM 2 ----- //
            var cellFirstTopDryer2Past = rowFirstTopDryer.Cells[4];
            cellFirstTopDryer2Past.Format.Font.Size = 6.5;
            cellFirstTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cellFirstTopDryer2Past.AddParagraph($"100");

            var cellFirstTopDryer2Current = rowFirstTopDryer.Cells[5];
            cellFirstTopDryer2Current.Format.Font.Size = 6.5;
            cellFirstTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryer2Age = 0;
            if (clothingFirstTopDryer2 != null)
            {
                firstTopDryer2Age = clothingFirstTopDryer2.Age != null ? Convert.ToInt32(clothingFirstTopDryer2.Age) : 0;
            }

            cellFirstTopDryer2Current.AddParagraph($"{firstTopDryer2Age}");

            var cellFirstTopDryer2Goal = rowFirstTopDryer.Cells[6];
            cellFirstTopDryer2Goal.Format.Font.Size = 6.5;
            cellFirstTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 6);
            cellFirstTopDryer2Goal.AddParagraph($"{firstTopDryer2Goal.Goal1}");

            var cellFirstTopDryer2Info = rowFirstTopDryer2.Cells[4];
            cellFirstTopDryer2Info.Format.Font.Size = 6.5;
            cellFirstTopDryer2Info.MergeRight = 2;
            cellFirstTopDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstTopDryer2 != null)
            {
                cellFirstTopDryer2Info.AddParagraph($"{clothingFirstTopDryer2.Manufacturer} {clothingFirstTopDryer2.Serial_Number}");
            }
            else cellFirstTopDryer2Info.AddParagraph("NA");
            #endregion

            #region 1st Top PM 3
            // ----- 1st Top Position PM 3 ----- //
            var cellFirstTopDryer3Past = rowFirstTopDryer.Cells[7];
            cellFirstTopDryer3Past.Format.Font.Size = 6.5;
            cellFirstTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cellFirstTopDryer3Past.AddParagraph($"100");

            var cellFirstTopDryer3Current = rowFirstTopDryer.Cells[8];
            cellFirstTopDryer3Current.Format.Font.Size = 6.5;
            cellFirstTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryer3Age = 0;
            if (clothingFirstTopDryer3 != null)
            {
                firstTopDryer3Age = clothingFirstTopDryer3.Age != null ? Convert.ToInt32(clothingFirstTopDryer3.Age) : 0;
            }

            cellFirstTopDryer3Current.AddParagraph($"{firstTopDryer3Age}");

            var cellFirstTopDryer3Goal = rowFirstTopDryer.Cells[9];
            cellFirstTopDryer3Goal.Format.Font.Size = 6.5;
            cellFirstTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 6);
            cellFirstTopDryer3Goal.AddParagraph($"{firstTopDryer3Goal.Goal1}");

            var cellFirstTopDryer3Info = rowFirstTopDryer2.Cells[7];
            cellFirstTopDryer3Info.Format.Font.Size = 6.5;
            cellFirstTopDryer3Info.MergeRight = 2;
            cellFirstTopDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstTopDryer3 != null)
            {
                cellFirstTopDryer3Info.AddParagraph($"{clothingFirstTopDryer3.Manufacturer} {clothingFirstTopDryer3.Serial_Number}");
            }
            else cellFirstTopDryer3Info.AddParagraph("NA");
            #endregion

            #region 1st Top PM 4
            // ----- 1st Top Position PM 4 ----- //
            var cellFirstTopDryer4Past = rowFirstTopDryer.Cells[10];
            cellFirstTopDryer4Past.Format.Font.Size = 6.5;
            cellFirstTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cellFirstTopDryer4Past.AddParagraph($"100");

            var cellFirstTopDryer4Current = rowFirstTopDryer.Cells[11];
            cellFirstTopDryer4Current.Format.Font.Size = 6.5;
            cellFirstTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryer4Age = 0;
            if (clothingFirstTopDryer4 != null)
            {
                firstTopDryer4Age = clothingFirstTopDryer4.Age != null ? Convert.ToInt32(clothingFirstTopDryer4.Age) : 0;
            }

            cellFirstTopDryer4Current.AddParagraph($"{firstTopDryer4Age}");

            var cellFirstTopDryer4Goal = rowFirstTopDryer.Cells[12];
            cellFirstTopDryer4Goal.Format.Font.Size = 6.5;
            cellFirstTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstTopDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 6);
            cellFirstTopDryer4Goal.AddParagraph($"{firstTopDryer4Goal.Goal1}");

            var cellFirstTopDryer4Info = rowFirstTopDryer2.Cells[10];
            cellFirstTopDryer4Info.Format.Font.Size = 6.5;
            cellFirstTopDryer4Info.MergeRight = 2;
            cellFirstTopDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothingFirstTopDryer4 != null)
            {
                cellFirstTopDryer4Info.AddParagraph($"{clothingFirstTopDryer4.Manufacturer} {clothingFirstTopDryer4.Serial_Number}");
            }
            else cellFirstTopDryer4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region 2nd Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 1);
            var clothing2ndTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 2);
            var clothing2ndTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 3);
            var clothing2ndTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 4);
            var row2ndTopDryer = weeklyPMTable.AddRow();
            row2ndTopDryer.Height = 10;
            var row2ndTopDryer2 = weeklyPMTable.AddRow();
            row2ndTopDryer2.Height = 10;

            //define 2nd Top position row cells

            var cell2ndTopDryer = row2ndTopDryer.Cells[0];
            cell2ndTopDryer.Format.Font.Size = 6.5;
            cell2ndTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer.MergeDown = 1;
            cell2ndTopDryer.AddParagraph("2ND TOP");

            #region 2nd Top PM 1
            // ----- 2nd Top Position PM 1 ----- //
            var cell2ndTopDryer1Past = row2ndTopDryer.Cells[1];
            cell2ndTopDryer1Past.Format.Font.Size = 6.5;
            cell2ndTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer1Past.AddParagraph($"100");

            var cell2ndTopDryer1Current = row2ndTopDryer.Cells[2];
            cell2ndTopDryer1Current.Format.Font.Size = 6.5;
            cell2ndTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryer1Age = 0;
            if (clothing2ndTopDryer1 != null)
            {
                secondTopDryer1Age = clothing2ndTopDryer1.Age != null ? Convert.ToInt32(clothing2ndTopDryer1.Age) : 0;
            }

            cell2ndTopDryer1Current.AddParagraph($"{secondTopDryer1Age}");

            var cell2ndTopDryer1Goal = row2ndTopDryer.Cells[3];
            cell2ndTopDryer1Goal.Format.Font.Size = 6.5;
            cell2ndTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 7);
            cell2ndTopDryer1Goal.AddParagraph($"{secondTopDryerlGoal.Goal1}");

            var cell2ndTopDryer1Info = row2ndTopDryer2.Cells[1];
            cell2ndTopDryer1Info.Format.Font.Size = 6.5;
            cell2ndTopDryer1Info.MergeRight = 2;
            cell2ndTopDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndTopDryer1 != null)
            {
                cell2ndTopDryer1Info.AddParagraph($"{clothing2ndTopDryer1.Manufacturer} {clothing2ndTopDryer1.Serial_Number}");
            }
            else cell2ndTopDryer1Info.AddParagraph("NA");
            #endregion

            #region 2nd Top PM 2
            // ----- 2nd Top Position PM 2 ----- //
            var cell2ndTopDryer2Past = row2ndTopDryer.Cells[4];
            cell2ndTopDryer2Past.Format.Font.Size = 6.5;
            cell2ndTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer2Past.AddParagraph($"100");

            var cell2ndTopDryer2Current = row2ndTopDryer.Cells[5];
            cell2ndTopDryer2Current.Format.Font.Size = 6.5;
            cell2ndTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryer2Age = 0;
            if (clothing2ndTopDryer2 != null)
            {
                secondTopDryer2Age = clothing2ndTopDryer2.Age != null ? Convert.ToInt32(clothing2ndTopDryer2.Age) : 0;
            }

            cell2ndTopDryer2Current.AddParagraph($"{secondTopDryer2Age}");

            var cell2ndTopDryer2Goal = row2ndTopDryer.Cells[6];
            cell2ndTopDryer2Goal.Format.Font.Size = 6.5;
            cell2ndTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 7);
            cell2ndTopDryer2Goal.AddParagraph($"{secondTopDryer2Goal.Goal1}");

            var cell2ndTopDryer2Info = row2ndTopDryer2.Cells[4];
            cell2ndTopDryer2Info.Format.Font.Size = 6.5;
            cell2ndTopDryer2Info.MergeRight = 2;
            cell2ndTopDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndTopDryer2 != null)
            {
                cell2ndTopDryer2Info.AddParagraph($"{clothing2ndTopDryer2.Manufacturer} {clothing2ndTopDryer2.Serial_Number}");
            }
            else cell2ndTopDryer2Info.AddParagraph("NA");
            #endregion

            #region 2nd Top PM 3
            // ----- 2nd Top Position PM 3 ----- //
            var cell2ndTopDryer3Past = row2ndTopDryer.Cells[7];
            cell2ndTopDryer3Past.Format.Font.Size = 6.5;
            cell2ndTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer3Past.AddParagraph($"100");

            var cell2ndTopDryer3Current = row2ndTopDryer.Cells[8];
            cell2ndTopDryer3Current.Format.Font.Size = 6.5;
            cell2ndTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryer3Age = 0;
            if (clothing2ndTopDryer3 != null)
            {
                secondTopDryer3Age = clothing2ndTopDryer3.Age != null ? Convert.ToInt32(clothing2ndTopDryer3.Age) : 0;
            }

            cell2ndTopDryer3Current.AddParagraph($"{secondTopDryer3Age}");

            var cell2ndTopDryer3Goal = row2ndTopDryer.Cells[9];
            cell2ndTopDryer3Goal.Format.Font.Size = 6.5;
            cell2ndTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 7);
            cell2ndTopDryer3Goal.AddParagraph($"{secondTopDryer3Goal.Goal1}");

            var cell2ndTopDryer3Info = row2ndTopDryer2.Cells[7];
            cell2ndTopDryer3Info.Format.Font.Size = 6.5;
            cell2ndTopDryer3Info.MergeRight = 2;
            cell2ndTopDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndTopDryer3 != null)
            {
                cell2ndTopDryer3Info.AddParagraph($"{clothing2ndTopDryer3.Manufacturer} {clothing2ndTopDryer3.Serial_Number}");
            }
            else cell2ndTopDryer3Info.AddParagraph("NA");
            #endregion

            #region 2nd Top PM 4
            // ----- 2nd Top Position PM 4 ----- //
            var cell2ndTopDryer4Past = row2ndTopDryer.Cells[10];
            cell2ndTopDryer4Past.Format.Font.Size = 6.5;
            cell2ndTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer4Past.AddParagraph($"100");

            var cell2ndTopDryer4Current = row2ndTopDryer.Cells[11];
            cell2ndTopDryer4Current.Format.Font.Size = 6.5;
            cell2ndTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryer4Age = 0;
            if (clothing2ndTopDryer4 != null)
            {
                secondTopDryer4Age = clothing2ndTopDryer4.Age != null ? Convert.ToInt32(clothing2ndTopDryer4.Age) : 0;
            }

            cell2ndTopDryer4Current.AddParagraph($"{secondTopDryer4Age}");

            var cell2ndTopDryer4Goal = row2ndTopDryer.Cells[12];
            cell2ndTopDryer4Goal.Format.Font.Size = 6.5;
            cell2ndTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondTopDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 7);
            cell2ndTopDryer4Goal.AddParagraph($"{secondTopDryer4Goal.Goal1}");

            var cell2ndTopDryer4Info = row2ndTopDryer2.Cells[10];
            cell2ndTopDryer4Info.Format.Font.Size = 6.5;
            cell2ndTopDryer4Info.MergeRight = 2;
            cell2ndTopDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndTopDryer4 != null)
            {
                cell2ndTopDryer4Info.AddParagraph($"{clothing2ndTopDryer4.Manufacturer} {clothing2ndTopDryer4.Serial_Number}");
            }
            else cell2ndTopDryer4Info.AddParagraph("NA");
            #endregion

            #endregion //2nd Top Dryer Position

            #region 3rd Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 1);
            var clothing3rdTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 2);
            var clothing3rdTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 3);
            var clothing3rdTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 4);
            var row3rdTopDryer = weeklyPMTable.AddRow();
            row3rdTopDryer.Height = 10;
            var row3rdTopDryer2 = weeklyPMTable.AddRow();
            row3rdTopDryer2.Height = 10;

            //define 3rd Top position row cells

            var cell3rdTopDryer = row3rdTopDryer.Cells[0];
            cell3rdTopDryer.Format.Font.Size = 6.5;
            cell3rdTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer.MergeDown = 1;
            cell3rdTopDryer.AddParagraph("3RD TOP");

            #region 3rd Top PM 1
            // ----- 3rd Top Position PM 1 ----- //
            var cell3rdTopDryer1Past = row3rdTopDryer.Cells[1];
            cell3rdTopDryer1Past.Format.Font.Size = 6.5;
            cell3rdTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer1Past.AddParagraph($"100");

            var cell3rdTopDryer1Current = row3rdTopDryer.Cells[2];
            cell3rdTopDryer1Current.Format.Font.Size = 6.5;
            cell3rdTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryer1Age = 0;
            if (clothing3rdTopDryer1 != null)
            {
                thirdTopDryer1Age = clothing3rdTopDryer1.Age != null ? Convert.ToInt32(clothing3rdTopDryer1.Age) : 0;
            }

            cell3rdTopDryer1Current.AddParagraph($"{thirdTopDryer1Age}");

            var cell3rdTopDryer1Goal = row3rdTopDryer.Cells[3];
            cell3rdTopDryer1Goal.Format.Font.Size = 6.5;
            cell3rdTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 8);
            cell3rdTopDryer1Goal.AddParagraph($"{thirdTopDryerlGoal.Goal1}");

            var cell3rdTopDryer1Info = row3rdTopDryer2.Cells[1];
            cell3rdTopDryer1Info.Format.Font.Size = 6.5;
            cell3rdTopDryer1Info.MergeRight = 2;
            cell3rdTopDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdTopDryer1 != null)
            {
                cell3rdTopDryer1Info.AddParagraph($"{clothing3rdTopDryer1.Manufacturer} {clothing3rdTopDryer1.Serial_Number}");
            }
            else cell3rdTopDryer1Info.AddParagraph("NA");
            #endregion

            #region 3rd Top PM 2
            // ----- 3rd Top Position PM 2 ----- //
            var cell3rdTopDryer2Past = row3rdTopDryer.Cells[4];
            cell3rdTopDryer2Past.Format.Font.Size = 6.5;
            cell3rdTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer2Past.AddParagraph($"100");

            var cell3rdTopDryer2Current = row3rdTopDryer.Cells[5];
            cell3rdTopDryer2Current.Format.Font.Size = 6.5;
            cell3rdTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryer2Age = 0;
            if (clothing3rdTopDryer2 != null)
            {
                thirdTopDryer2Age = clothing3rdTopDryer2.Age != null ? Convert.ToInt32(clothing3rdTopDryer2.Age) : 0;
            }

            cell3rdTopDryer2Current.AddParagraph($"{thirdTopDryer2Age}");

            var cell3rdTopDryer2Goal = row3rdTopDryer.Cells[6];
            cell3rdTopDryer2Goal.Format.Font.Size = 6.5;
            cell3rdTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 8);
            cell3rdTopDryer2Goal.AddParagraph($"{thirdTopDryer2Goal.Goal1}");

            var cell3rdTopDryer2Info = row3rdTopDryer2.Cells[4];
            cell3rdTopDryer2Info.Format.Font.Size = 6.5;
            cell3rdTopDryer2Info.MergeRight = 2;
            cell3rdTopDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdTopDryer2 != null)
            {
                cell3rdTopDryer2Info.AddParagraph($"{clothing3rdTopDryer2.Manufacturer} {clothing3rdTopDryer2.Serial_Number}");
            }
            else cell3rdTopDryer2Info.AddParagraph("NA");
            #endregion

            #region 3rd Top PM 3

            #endregion

            #region 3rd Top PM 4

            #endregion

            #endregion // 3rd Top Dryer Position

            #region 4th Top Dryer Position

            #endregion

            #region 1st Bottom Dryer Position

            #endregion

            #region 2nd Bottom Dryer Position

            #endregion

            #region 3rd Bottom Dryer Position

            #endregion

            #region 4th Bottom Dryer Position

            #endregion

            #region Breast Roll Position

            #endregion

            #region Dandy Position

            #endregion

            #region Lumpbreaker Position

            #endregion

            #region Suction Pickup Position

            #endregion

            #region 1st Press Top Position

            #endregion

            #region 1st Press Bottom Position

            #endregion

            #region 2nd Press Top Position

            #endregion

            #region 2nd Press Bottom Position

            #endregion

            #region 3rd Press Top Position

            #endregion

            #region 3rd Press Bottom Position

            #endregion

            #region Smooth Top Position

            #endregion

            #region Smooth Bottom Position

            #endregion

            #region Hard Size Press Position

            #endregion

            #region Soft Size Press Position

            #endregion

            #region Aquitherm Top Position

            #endregion

            #region Nibco Bottom Position

            #endregion

            return weeklyPMTable;
        }
    }
}