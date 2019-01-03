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
            // ----- 3rd Top Position PM 3 ----- //
            var cell3rdTopDryer3Past = row3rdTopDryer.Cells[7];
            cell3rdTopDryer3Past.Format.Font.Size = 6.5;
            cell3rdTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer3Past.AddParagraph($"100");

            var cell3rdTopDryer3Current = row3rdTopDryer.Cells[8];
            cell3rdTopDryer3Current.Format.Font.Size = 6.5;
            cell3rdTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryer3Age = 0;
            if (clothing3rdTopDryer3 != null)
            {
                thirdTopDryer3Age = clothing3rdTopDryer3.Age != null ? Convert.ToInt32(clothing3rdTopDryer3.Age) : 0;
            }

            cell3rdTopDryer3Current.AddParagraph($"{thirdTopDryer3Age}");

            var cell3rdTopDryer3Goal = row3rdTopDryer.Cells[9];
            cell3rdTopDryer3Goal.Format.Font.Size = 6.5;
            cell3rdTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 8);
            cell3rdTopDryer3Goal.AddParagraph($"{thirdTopDryer3Goal.Goal1}");

            var cell3rdTopDryer3Info = row3rdTopDryer2.Cells[7];
            cell3rdTopDryer3Info.Format.Font.Size = 6.5;
            cell3rdTopDryer3Info.MergeRight = 2;
            cell3rdTopDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdTopDryer3 != null)
            {
                cell3rdTopDryer3Info.AddParagraph($"{clothing3rdTopDryer3.Manufacturer} {clothing3rdTopDryer3.Serial_Number}");
            }
            else cell3rdTopDryer3Info.AddParagraph("NA");
            #endregion

            #region 3rd Top PM 4
            // ----- 3rd Top Position PM 4 ----- //
            var cell3rdTopDryer4Past = row3rdTopDryer.Cells[10];
            cell3rdTopDryer4Past.Format.Font.Size = 6.5;
            cell3rdTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer4Past.AddParagraph($"100");

            var cell3rdTopDryer4Current = row3rdTopDryer.Cells[11];
            cell3rdTopDryer4Current.Format.Font.Size = 6.5;
            cell3rdTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryer4Age = 0;
            if (clothing3rdTopDryer4 != null)
            {
                thirdTopDryer4Age = clothing3rdTopDryer4.Age != null ? Convert.ToInt32(clothing3rdTopDryer4.Age) : 0;
            }

            cell3rdTopDryer4Current.AddParagraph($"{thirdTopDryer4Age}");

            var cell3rdTopDryer4Goal = row3rdTopDryer.Cells[12];
            cell3rdTopDryer4Goal.Format.Font.Size = 6.5;
            cell3rdTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdTopDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 8);
            cell3rdTopDryer4Goal.AddParagraph($"{thirdTopDryer4Goal.Goal1}");

            var cell3rdTopDryer4Info = row3rdTopDryer2.Cells[10];
            cell3rdTopDryer4Info.Format.Font.Size = 6.5;
            cell3rdTopDryer4Info.MergeRight = 2;
            cell3rdTopDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdTopDryer4 != null)
            {
                cell3rdTopDryer4Info.AddParagraph($"{clothing3rdTopDryer4.Manufacturer} {clothing3rdTopDryer4.Serial_Number}");
            }
            else cell3rdTopDryer4Info.AddParagraph("NA");
            #endregion

            #endregion // 3rd Top Dryer Position

            #region 4th Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing4thTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 1);
            var clothing4thTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 2);
            var clothing4thTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 3);
            var clothing4thTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 4);
            var row4thTopDryer = weeklyPMTable.AddRow();
            row4thTopDryer.Height = 10;
            var row4thTopDryer2 = weeklyPMTable.AddRow();
            row4thTopDryer2.Height = 10;

            //define 4th Top position row cells

            var cell4thTopDryer = row4thTopDryer.Cells[0];
            cell4thTopDryer.Format.Font.Size = 6.5;
            cell4thTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer.MergeDown = 1;
            cell4thTopDryer.AddParagraph("4TH TOP");

            #region 4th Top PM 1
            // ----- 4th Top Position PM 1 ----- //
            var cell4thTopDryer1Past = row4thTopDryer.Cells[1];
            cell4thTopDryer1Past.Format.Font.Size = 6.5;
            cell4thTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer1Past.AddParagraph($"100");

            var cell4thTopDryer1Current = row4thTopDryer.Cells[2];
            cell4thTopDryer1Current.Format.Font.Size = 6.5;
            cell4thTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryer1Age = 0;
            if (clothing4thTopDryer1 != null)
            {
                fourthTopDryer1Age = clothing4thTopDryer1.Age != null ? Convert.ToInt32(clothing4thTopDryer1.Age) : 0;
            }

            cell4thTopDryer1Current.AddParagraph($"{fourthTopDryer1Age}");

            var cell4thTopDryer1Goal = row4thTopDryer.Cells[3];
            cell4thTopDryer1Goal.Format.Font.Size = 6.5;
            cell4thTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 9);
            cell4thTopDryer1Goal.AddParagraph($"{fourthTopDryerlGoal.Goal1}");

            var cell4thTopDryer1Info = row4thTopDryer2.Cells[1];
            cell4thTopDryer1Info.Format.Font.Size = 6.5;
            cell4thTopDryer1Info.MergeRight = 2;
            cell4thTopDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothing4thTopDryer1 != null)
            {
                cell4thTopDryer1Info.AddParagraph($"{clothing4thTopDryer1.Manufacturer} {clothing4thTopDryer1.Serial_Number}");
            }
            else cell4thTopDryer1Info.AddParagraph("NA");
            #endregion

            #region 4th Top PM 2
            // ----- 4th Top Position PM 2 ----- //
            var cell4thTopDryer2Past = row4thTopDryer.Cells[4];
            cell4thTopDryer2Past.Format.Font.Size = 6.5;
            cell4thTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer2Past.AddParagraph($"100");

            var cell4thTopDryer2Current = row4thTopDryer.Cells[5];
            cell4thTopDryer2Current.Format.Font.Size = 6.5;
            cell4thTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryer2Age = 0;
            if (clothing4thTopDryer2 != null)
            {
                fourthTopDryer2Age = clothing4thTopDryer2.Age != null ? Convert.ToInt32(clothing4thTopDryer2.Age) : 0;
            }

            cell4thTopDryer2Current.AddParagraph($"{fourthTopDryer2Age}");

            var cell4thTopDryer2Goal = row4thTopDryer.Cells[6];
            cell4thTopDryer2Goal.Format.Font.Size = 6.5;
            cell4thTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 9);
            cell4thTopDryer2Goal.AddParagraph($"{fourthTopDryer2Goal.Goal1}");

            var cell4thTopDryer2Info = row4thTopDryer2.Cells[4];
            cell4thTopDryer2Info.Format.Font.Size = 6.5;
            cell4thTopDryer2Info.MergeRight = 2;
            cell4thTopDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothing4thTopDryer2 != null)
            {
                cell4thTopDryer2Info.AddParagraph($"{clothing4thTopDryer2.Manufacturer} {clothing4thTopDryer2.Serial_Number}");
            }
            else cell4thTopDryer2Info.AddParagraph("NA");

            #endregion

            #region 4th Top PM 3
            // ----- 4th Top Position PM 3 ----- //
            var cell4thTopDryer3Past = row4thTopDryer.Cells[7];
            cell4thTopDryer3Past.Format.Font.Size = 6.5;
            cell4thTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer3Past.AddParagraph($"100");

            var cell4thTopDryer3Current = row4thTopDryer.Cells[8];
            cell4thTopDryer3Current.Format.Font.Size = 6.5;
            cell4thTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryer3Age = 0;
            if (clothing4thTopDryer3 != null)
            {
                fourthTopDryer3Age = clothing4thTopDryer3.Age != null ? Convert.ToInt32(clothing4thTopDryer3.Age) : 0;
            }

            cell4thTopDryer3Current.AddParagraph($"{fourthTopDryer3Age}");

            var cell4thTopDryer3Goal = row4thTopDryer.Cells[9];
            cell4thTopDryer3Goal.Format.Font.Size = 6.5;
            cell4thTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 9);
            cell4thTopDryer3Goal.AddParagraph($"{fourthTopDryer3Goal.Goal1}");

            var cell4thTopDryer3Info = row4thTopDryer2.Cells[7];
            cell4thTopDryer3Info.Format.Font.Size = 6.5;
            cell4thTopDryer3Info.MergeRight = 2;
            cell4thTopDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothing4thTopDryer3 != null)
            {
                cell4thTopDryer3Info.AddParagraph($"{clothing4thTopDryer3.Manufacturer} {clothing4thTopDryer3.Serial_Number}");
            }
            else cell4thTopDryer3Info.AddParagraph("NA");
            #endregion

            #region 4th Top PM 4
            // ----- 4th Top Position PM 4 ----- //
            var cell4thTopDryer4Past = row4thTopDryer.Cells[10];
            cell4thTopDryer4Past.Format.Font.Size = 6.5;
            cell4thTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer4Past.AddParagraph($"100");

            var cell4thTopDryer4Current = row4thTopDryer.Cells[11];
            cell4thTopDryer4Current.Format.Font.Size = 6.5;
            cell4thTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryer4Age = 0;
            if (clothing4thTopDryer4 != null)
            {
                fourthTopDryer4Age = clothing4thTopDryer4.Age != null ? Convert.ToInt32(clothing4thTopDryer4.Age) : 0;
            }

            cell4thTopDryer4Current.AddParagraph($"{fourthTopDryer4Age}");

            var cell4thTopDryer4Goal = row4thTopDryer.Cells[12];
            cell4thTopDryer4Goal.Format.Font.Size = 6.5;
            cell4thTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthTopDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 9);
            cell4thTopDryer4Goal.AddParagraph($"{fourthTopDryer4Goal.Goal1}");

            var cell4thTopDryer4Info = row4thTopDryer2.Cells[10];
            cell4thTopDryer4Info.Format.Font.Size = 6.5;
            cell4thTopDryer4Info.MergeRight = 2;
            cell4thTopDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothing4thTopDryer4 != null)
            {
                cell4thTopDryer4Info.AddParagraph($"{clothing4thTopDryer4.Manufacturer} {clothing4thTopDryer4.Serial_Number}");
            }
            else cell4thTopDryer4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region 1st Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 1);
            var clothing1stBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 2);
            var clothing1stBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 3);
            var clothing1stBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 4);
            var row1stBottomDryer = weeklyPMTable.AddRow();
            row1stBottomDryer.Height = 10;
            var row1stBottomDryer2 = weeklyPMTable.AddRow();
            row1stBottomDryer2.Height = 10;

            //define 1st Bottom position row cells

            var cell1stBottomDryer = row1stBottomDryer.Cells[0];
            cell1stBottomDryer.Format.Font.Size = 6.5;
            cell1stBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer.MergeDown = 1;
            cell1stBottomDryer.AddParagraph("1ST BOTTOM");

            #region 1st Bottom PM 1
            // ----- 1st Bottom Position PM 1 ----- //
            var cell1stBottomDryer1Past = row1stBottomDryer.Cells[1];
            cell1stBottomDryer1Past.Format.Font.Size = 6.5;
            cell1stBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer1Past.AddParagraph($"100");

            var cell1stBottomDryer1Current = row1stBottomDryer.Cells[2];
            cell1stBottomDryer1Current.Format.Font.Size = 6.5;
            cell1stBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryer1Age = 0;
            if (clothing1stBottomDryer1 != null)
            {
                firstBottomDryer1Age = clothing1stBottomDryer1.Age != null ? Convert.ToInt32(clothing1stBottomDryer1.Age) : 0;
            }

            cell1stBottomDryer1Current.AddParagraph($"{firstBottomDryer1Age}");

            var cell1stBottomDryer1Goal = row1stBottomDryer.Cells[3];
            cell1stBottomDryer1Goal.Format.Font.Size = 6.5;
            cell1stBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 10);
            cell1stBottomDryer1Goal.AddParagraph($"{firstBottomDryerlGoal.Goal1}");

            var cell1stBottomDryer1Info = row1stBottomDryer2.Cells[1];
            cell1stBottomDryer1Info.Format.Font.Size = 6.5;
            cell1stBottomDryer1Info.MergeRight = 2;
            cell1stBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothing1stBottomDryer1 != null)
            {
                cell1stBottomDryer1Info.AddParagraph($"{clothing1stBottomDryer1.Manufacturer} {clothing1stBottomDryer1.Serial_Number}");
            }
            else cell1stBottomDryer1Info.AddParagraph("NA");
            #endregion

            #region 1st Bottom PM 2
            // ----- 1st Bottom Position PM 2 ----- //
            var cell1stBottomDryer2Past = row1stBottomDryer.Cells[4];
            cell1stBottomDryer2Past.Format.Font.Size = 6.5;
            cell1stBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer2Past.AddParagraph($"100");

            var cell1stBottomDryer2Current = row1stBottomDryer.Cells[5];
            cell1stBottomDryer2Current.Format.Font.Size = 6.5;
            cell1stBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryer2Age = 0;
            if (clothing1stBottomDryer2 != null)
            {
                firstBottomDryer2Age = clothing1stBottomDryer2.Age != null ? Convert.ToInt32(clothing1stBottomDryer2.Age) : 0;
            }

            cell1stBottomDryer2Current.AddParagraph($"{firstBottomDryer2Age}");

            var cell1stBottomDryer2Goal = row1stBottomDryer.Cells[6];
            cell1stBottomDryer2Goal.Format.Font.Size = 6.5;
            cell1stBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 10);
            cell1stBottomDryer2Goal.AddParagraph($"{firstBottomDryer2Goal.Goal1}");

            var cell1stBottomDryer2Info = row1stBottomDryer2.Cells[4];
            cell1stBottomDryer2Info.Format.Font.Size = 6.5;
            cell1stBottomDryer2Info.MergeRight = 2;
            cell1stBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothing1stBottomDryer2 != null)
            {
                cell1stBottomDryer2Info.AddParagraph($"{clothing1stBottomDryer2.Manufacturer} {clothing1stBottomDryer2.Serial_Number}");
            }
            else cell1stBottomDryer2Info.AddParagraph("NA");
            #endregion

            #region 1st Bottom PM 3
            // ----- 1st Bottom Position PM 3 ----- //
            var cell1stBottomDryer3Past = row1stBottomDryer.Cells[7];
            cell1stBottomDryer3Past.Format.Font.Size = 6.5;
            cell1stBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer3Past.AddParagraph($"100");

            var cell1stBottomDryer3Current = row1stBottomDryer.Cells[8];
            cell1stBottomDryer3Current.Format.Font.Size = 6.5;
            cell1stBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryer3Age = 0;
            if (clothing1stBottomDryer3 != null)
            {
                firstBottomDryer3Age = clothing1stBottomDryer3.Age != null ? Convert.ToInt32(clothing1stBottomDryer3.Age) : 0;
            }

            cell1stBottomDryer3Current.AddParagraph($"{firstBottomDryer3Age}");

            var cell1stBottomDryer3Goal = row1stBottomDryer.Cells[9];
            cell1stBottomDryer3Goal.Format.Font.Size = 6.5;
            cell1stBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 10);
            cell1stBottomDryer3Goal.AddParagraph($"{firstBottomDryer3Goal.Goal1}");

            var cell1stBottomDryer3Info = row1stBottomDryer2.Cells[7];
            cell1stBottomDryer3Info.Format.Font.Size = 6.5;
            cell1stBottomDryer3Info.MergeRight = 2;
            cell1stBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothing1stBottomDryer3 != null)
            {
                cell1stBottomDryer3Info.AddParagraph($"{clothing1stBottomDryer3.Manufacturer} {clothing1stBottomDryer3.Serial_Number}");
            }
            else cell1stBottomDryer3Info.AddParagraph("NA");
            #endregion

            #region 1st Bottom PM 4
            // ----- 1st Bottom Position PM 4 ----- //
            var cell1stBottomDryer4Past = row1stBottomDryer.Cells[10];
            cell1stBottomDryer4Past.Format.Font.Size = 6.5;
            cell1stBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer4Past.AddParagraph($"100");

            var cell1stBottomDryer4Current = row1stBottomDryer.Cells[11];
            cell1stBottomDryer4Current.Format.Font.Size = 6.5;
            cell1stBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryer4Age = 0;
            if (clothing1stBottomDryer4 != null)
            {
                firstBottomDryer4Age = clothing1stBottomDryer4.Age != null ? Convert.ToInt32(clothing1stBottomDryer4.Age) : 0;
            }

            cell1stBottomDryer4Current.AddParagraph($"{firstBottomDryer4Age}");

            var cell1stBottomDryer4Goal = row1stBottomDryer.Cells[12];
            cell1stBottomDryer4Goal.Format.Font.Size = 6.5;
            cell1stBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstBottomDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 10);
            cell1stBottomDryer4Goal.AddParagraph($"{firstBottomDryer4Goal.Goal1}");

            var cell1stBottomDryer4Info = row1stBottomDryer2.Cells[10];
            cell1stBottomDryer4Info.Format.Font.Size = 6.5;
            cell1stBottomDryer4Info.MergeRight = 2;
            cell1stBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothing1stBottomDryer4 != null)
            {
                cell1stBottomDryer4Info.AddParagraph($"{clothing1stBottomDryer4.Manufacturer} {clothing1stBottomDryer4.Serial_Number}");
            }
            else cell1stBottomDryer4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region 2nd Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 1);
            var clothing2ndBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 2);
            var clothing2ndBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 3);
            var clothing2ndBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 4);
            var row2ndBottomDryer = weeklyPMTable.AddRow();
            row2ndBottomDryer.Height = 10;
            var row2ndBottomDryer2 = weeklyPMTable.AddRow();
            row2ndBottomDryer2.Height = 10;

            //define 1st Bottom position row cells

            var cell2ndBottomDryer = row2ndBottomDryer.Cells[0];
            cell2ndBottomDryer.Format.Font.Size = 6.5;
            cell2ndBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer.MergeDown = 1;
            cell2ndBottomDryer.AddParagraph("2ND BOTTOM");

            #region 2nd Bottom PM 1
            // ----- 2nd Bottom Position PM 1 ----- //
            var cell2ndBottomDryer1Past = row2ndBottomDryer.Cells[1];
            cell2ndBottomDryer1Past.Format.Font.Size = 6.5;
            cell2ndBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer1Past.AddParagraph($"100");

            var cell2ndtBottomDryer1Current = row2ndBottomDryer.Cells[2];
            cell2ndtBottomDryer1Current.Format.Font.Size = 6.5;
            cell2ndtBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var secondtBottomDryer1Age = 0;
            if (clothing2ndBottomDryer1 != null)
            {
                secondtBottomDryer1Age = clothing2ndBottomDryer1.Age != null ? Convert.ToInt32(clothing2ndBottomDryer1.Age) : 0;
            }

            cell2ndtBottomDryer1Current.AddParagraph($"{secondtBottomDryer1Age}");

            var cell2ndBottomDryer1Goal = row2ndBottomDryer.Cells[3];
            cell2ndBottomDryer1Goal.Format.Font.Size = 6.5;
            cell2ndBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondBottomDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 11);
            cell2ndBottomDryer1Goal.AddParagraph($"{secondBottomDryerlGoal.Goal1}");

            var cell2ndBottomDryer1Info = row2ndBottomDryer2.Cells[1];
            cell2ndBottomDryer1Info.Format.Font.Size = 6.5;
            cell2ndBottomDryer1Info.MergeRight = 2;
            cell2ndBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndBottomDryer1 != null)
            {
                cell2ndBottomDryer1Info.AddParagraph($"{clothing2ndBottomDryer1.Manufacturer} {clothing2ndBottomDryer1.Serial_Number}");
            }
            else cell2ndBottomDryer1Info.AddParagraph("NA");
            #endregion

            #region 2nd Bottom PM 2
            // ----- 2nd Bottom Position PM 2 ----- //
            var cell2ndBottomDryer2Past = row2ndBottomDryer.Cells[4];
            cell2ndBottomDryer2Past.Format.Font.Size = 6.5;
            cell2ndBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer2Past.AddParagraph($"100");

            var cell2ndBottomDryer2Current = row2ndBottomDryer.Cells[5];
            cell2ndBottomDryer2Current.Format.Font.Size = 6.5;
            cell2ndBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var secondtBottomDryer2Age = 0;
            if (clothing2ndBottomDryer2 != null)
            {
                secondtBottomDryer2Age = clothing2ndBottomDryer2.Age != null ? Convert.ToInt32(clothing2ndBottomDryer2.Age) : 0;
            }

            cell2ndBottomDryer2Current.AddParagraph($"{secondtBottomDryer2Age}");

            var cell2ndBottomDryer2Goal = row2ndBottomDryer.Cells[6];
            cell2ndBottomDryer2Goal.Format.Font.Size = 6.5;
            cell2ndBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondBottomDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 11);
            cell2ndBottomDryer2Goal.AddParagraph($"{secondBottomDryer2Goal.Goal1}");

            var cell2ndBottomDryer2Info = row2ndBottomDryer2.Cells[4];
            cell2ndBottomDryer2Info.Format.Font.Size = 6.5;
            cell2ndBottomDryer2Info.MergeRight = 2;
            cell2ndBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndBottomDryer2 != null)
            {
                cell2ndBottomDryer2Info.AddParagraph($"{clothing2ndBottomDryer2.Manufacturer} {clothing2ndBottomDryer2.Serial_Number}");
            }
            else cell2ndBottomDryer2Info.AddParagraph("NA");
            #endregion

            #region 2nd Bottom PM 3
            // ----- 2nd Bottom Position PM 3 ----- //
            var cell2ndBottomDryer3Past = row2ndBottomDryer.Cells[7];
            cell2ndBottomDryer3Past.Format.Font.Size = 6.5;
            cell2ndBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer3Past.AddParagraph($"100");

            var cell2ndBottomDryer3Current = row2ndBottomDryer.Cells[8];
            cell2ndBottomDryer3Current.Format.Font.Size = 6.5;
            cell2ndBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var secondtBottomDryer3Age = 0;
            if (clothing2ndBottomDryer3 != null)
            {
                secondtBottomDryer3Age = clothing2ndBottomDryer3.Age != null ? Convert.ToInt32(clothing2ndBottomDryer3.Age) : 0;
            }

            cell2ndBottomDryer3Current.AddParagraph($"{secondtBottomDryer3Age}");

            var cell2ndBottomDryer3Goal = row2ndBottomDryer.Cells[9];
            cell2ndBottomDryer3Goal.Format.Font.Size = 6.5;
            cell2ndBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondBottomDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 11);
            cell2ndBottomDryer3Goal.AddParagraph($"{secondBottomDryer3Goal.Goal1}");

            var cell2ndBottomDryer3Info = row2ndBottomDryer2.Cells[7];
            cell2ndBottomDryer3Info.Format.Font.Size = 6.5;
            cell2ndBottomDryer3Info.MergeRight = 2;
            cell2ndBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndBottomDryer3 != null)
            {
                cell2ndBottomDryer3Info.AddParagraph($"{clothing2ndBottomDryer3.Manufacturer} {clothing2ndBottomDryer3.Serial_Number}");
            }
            else cell2ndBottomDryer3Info.AddParagraph("NA");
            #endregion

            #region 2nd Bottom PM 4
            // ----- 2nd Bottom Position PM 4 ----- //
            var cell2ndBottomDryer4Past = row2ndBottomDryer.Cells[10];
            cell2ndBottomDryer4Past.Format.Font.Size = 6.5;
            cell2ndBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer4Past.AddParagraph($"100");

            var cell2ndBottomDryer4Current = row2ndBottomDryer.Cells[11];
            cell2ndBottomDryer4Current.Format.Font.Size = 6.5;
            cell2ndBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var secondtBottomDryer4Age = 0;
            if (clothing2ndBottomDryer4 != null)
            {
                secondtBottomDryer4Age = clothing2ndBottomDryer4.Age != null ? Convert.ToInt32(clothing2ndBottomDryer4.Age) : 0;
            }

            cell2ndBottomDryer4Current.AddParagraph($"{secondtBottomDryer4Age}");

            var cell2ndBottomDryer4Goal = row2ndBottomDryer.Cells[12];
            cell2ndBottomDryer4Goal.Format.Font.Size = 6.5;
            cell2ndBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondBottomDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 11);
            cell2ndBottomDryer4Goal.AddParagraph($"{secondBottomDryer4Goal.Goal1}");

            var cell2ndBottomDryer4Info = row2ndBottomDryer2.Cells[10];
            cell2ndBottomDryer4Info.Format.Font.Size = 6.5;
            cell2ndBottomDryer4Info.MergeRight = 2;
            cell2ndBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndBottomDryer4 != null)
            {
                cell2ndBottomDryer4Info.AddParagraph($"{clothing2ndBottomDryer4.Manufacturer} {clothing2ndBottomDryer4.Serial_Number}");
            }
            else cell2ndBottomDryer4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region 3rd Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 1);
            var clothing3rdBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 2);
            var clothing3rdBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 3);
            var clothing3rdBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 4);
            var row3rdBottomDryer = weeklyPMTable.AddRow();
            row3rdBottomDryer.Height = 10;
            var row3rdBottomDryer2 = weeklyPMTable.AddRow();
            row3rdBottomDryer2.Height = 10;

            //define 3rd Bottom position row cells

            var cell3rdBottomDryer = row3rdBottomDryer.Cells[0];
            cell3rdBottomDryer.Format.Font.Size = 6.5;
            cell3rdBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer.MergeDown = 1;
            cell3rdBottomDryer.AddParagraph("3RD BOTTOM");

            #region 3rd Bottom PM 1
            // ----- 3rd Bottom Position PM 1 ----- //
            var cell3rdBottomDryer1Past = row3rdBottomDryer.Cells[1];
            cell3rdBottomDryer1Past.Format.Font.Size = 6.5;
            cell3rdBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer1Past.AddParagraph($"100");

            var cell3rdBottomDryer1Current = row3rdBottomDryer.Cells[2];
            cell3rdBottomDryer1Current.Format.Font.Size = 6.5;
            cell3rdBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdtBottomDryer1Age = 0;
            if (clothing3rdBottomDryer1 != null)
            {
                thirdtBottomDryer1Age = clothing3rdBottomDryer1.Age != null ? Convert.ToInt32(clothing3rdBottomDryer1.Age) : 0;
            }

            cell3rdBottomDryer1Current.AddParagraph($"{thirdtBottomDryer1Age}");

            var cell3rdBottomDryer1Goal = row3rdBottomDryer.Cells[3];
            cell3rdBottomDryer1Goal.Format.Font.Size = 6.5;
            cell3rdBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdBottomDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 12);
            cell3rdBottomDryer1Goal.AddParagraph($"{thirdBottomDryerlGoal.Goal1}");

            var cell3rdBottomDryer1Info = row3rdBottomDryer2.Cells[1];
            cell3rdBottomDryer1Info.Format.Font.Size = 6.5;
            cell3rdBottomDryer1Info.MergeRight = 2;
            cell3rdBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdBottomDryer1 != null)
            {
                cell3rdBottomDryer1Info.AddParagraph($"{clothing3rdBottomDryer1.Manufacturer} {clothing3rdBottomDryer1.Serial_Number}");
            }
            else cell3rdBottomDryer1Info.AddParagraph("NA");
            #endregion

            #region 3rd Bottom PM 2
            // ----- 3rd Bottom Position PM 2 ----- //
            var cell3rdBottomDryer2Past = row3rdBottomDryer.Cells[4];
            cell3rdBottomDryer2Past.Format.Font.Size = 6.5;
            cell3rdBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer2Past.AddParagraph($"100");

            var cell3rdBottomDryer2Current = row3rdBottomDryer.Cells[5];
            cell3rdBottomDryer2Current.Format.Font.Size = 6.5;
            cell3rdBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdtBottomDryer2Age = 0;
            if (clothing3rdBottomDryer2 != null)
            {
                thirdtBottomDryer2Age = clothing3rdBottomDryer2.Age != null ? Convert.ToInt32(clothing3rdBottomDryer2.Age) : 0;
            }

            cell3rdBottomDryer2Current.AddParagraph($"{thirdtBottomDryer2Age}");

            var cell3rdBottomDryer2Goal = row3rdBottomDryer.Cells[6];
            cell3rdBottomDryer2Goal.Format.Font.Size = 6.5;
            cell3rdBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdBottomDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 12);
            cell3rdBottomDryer2Goal.AddParagraph($"{thirdBottomDryer2Goal.Goal1}");

            var cell3rdBottomDryer2Info = row3rdBottomDryer2.Cells[4];
            cell3rdBottomDryer2Info.Format.Font.Size = 6.5;
            cell3rdBottomDryer2Info.MergeRight = 2;
            cell3rdBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdBottomDryer2 != null)
            {
                cell3rdBottomDryer2Info.AddParagraph($"{clothing3rdBottomDryer2.Manufacturer} {clothing3rdBottomDryer2.Serial_Number}");
            }
            else cell3rdBottomDryer2Info.AddParagraph("NA");
            #endregion

            #region 3rd Bottom PM 3
            // ----- 3rd Bottom Position PM 3 ----- //
            var cell3rdBottomDryer3Past = row3rdBottomDryer.Cells[7];
            cell3rdBottomDryer3Past.Format.Font.Size = 6.5;
            cell3rdBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer3Past.AddParagraph($"100");

            var cell3rdBottomDryer3Current = row3rdBottomDryer.Cells[8];
            cell3rdBottomDryer3Current.Format.Font.Size = 6.5;
            cell3rdBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdtBottomDryer3Age = 0;
            if (clothing3rdBottomDryer3 != null)
            {
                thirdtBottomDryer3Age = clothing3rdBottomDryer3.Age != null ? Convert.ToInt32(clothing3rdBottomDryer3.Age) : 0;
            }

            cell3rdBottomDryer3Current.AddParagraph($"{thirdtBottomDryer3Age}");

            var cell3rdBottomDryer3Goal = row3rdBottomDryer.Cells[9];
            cell3rdBottomDryer3Goal.Format.Font.Size = 6.5;
            cell3rdBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdBottomDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 12);
            cell3rdBottomDryer3Goal.AddParagraph($"{thirdBottomDryer3Goal.Goal1}");

            var cell3rdBottomDryer3Info = row3rdBottomDryer2.Cells[7];
            cell3rdBottomDryer3Info.Format.Font.Size = 6.5;
            cell3rdBottomDryer3Info.MergeRight = 2;
            cell3rdBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdBottomDryer3 != null)
            {
                cell3rdBottomDryer3Info.AddParagraph($"{clothing3rdBottomDryer3.Manufacturer} {clothing3rdBottomDryer3.Serial_Number}");
            }
            else cell3rdBottomDryer3Info.AddParagraph("NA");
            #endregion

            #region 3rd Bottom PM 4
            // ----- 3rd Bottom Position PM 4 ----- //
            var cell3rdBottomDryer4Past = row3rdBottomDryer.Cells[10];
            cell3rdBottomDryer4Past.Format.Font.Size = 6.5;
            cell3rdBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer4Past.AddParagraph($"100");

            var cell3rdBottomDryer4Current = row3rdBottomDryer.Cells[11];
            cell3rdBottomDryer4Current.Format.Font.Size = 6.5;
            cell3rdBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdtBottomDryer4Age = 0;
            if (clothing3rdBottomDryer4 != null)
            {
                thirdtBottomDryer4Age = clothing3rdBottomDryer4.Age != null ? Convert.ToInt32(clothing3rdBottomDryer4.Age) : 0;
            }

            cell3rdBottomDryer4Current.AddParagraph($"{thirdtBottomDryer4Age}");

            var cell3rdBottomDryer4Goal = row3rdBottomDryer.Cells[12];
            cell3rdBottomDryer4Goal.Format.Font.Size = 6.5;
            cell3rdBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdBottomDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 12);
            cell3rdBottomDryer4Goal.AddParagraph($"{thirdBottomDryer4Goal.Goal1}");

            var cell3rdBottomDryer4Info = row3rdBottomDryer2.Cells[10];
            cell3rdBottomDryer4Info.Format.Font.Size = 6.5;
            cell3rdBottomDryer4Info.MergeRight = 2;
            cell3rdBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdBottomDryer4 != null)
            {
                cell3rdBottomDryer4Info.AddParagraph($"{clothing3rdBottomDryer4.Manufacturer} {clothing3rdBottomDryer4.Serial_Number}");
            }
            else cell3rdBottomDryer4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region 4th Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing4thBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 1);
            var clothing4thBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 2);
            var clothing4thBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 3);
            var clothing4thBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 4);
            var row4thBottomDryer = weeklyPMTable.AddRow();
            row4thBottomDryer.Height = 10;
            var row4thBottomDryer2 = weeklyPMTable.AddRow();
            row4thBottomDryer2.Height = 10;

            //define 3rd Bottom position row cells

            var cell4thBottomDryer = row4thBottomDryer.Cells[0];
            cell4thBottomDryer.Format.Font.Size = 6.5;
            cell4thBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer.MergeDown = 1;
            cell4thBottomDryer.AddParagraph("4TH BOTTOM");

            #region 4th Bottom PM 1
            // ----- 4th Bottom Position PM 1 ----- //
            var cell4thBottomDryer1Past = row4thBottomDryer.Cells[1];
            cell4thBottomDryer1Past.Format.Font.Size = 6.5;
            cell4thBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer1Past.AddParagraph($"100");

            var cell4thBottomDryer1Current = row4thBottomDryer.Cells[2];
            cell4thBottomDryer1Current.Format.Font.Size = 6.5;
            cell4thBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryer1Age = 0;
            if (clothing4thBottomDryer1 != null)
            {
                fourthBottomDryer1Age = clothing4thBottomDryer1.Age != null ? Convert.ToInt32(clothing4thBottomDryer1.Age) : 0;
            }

            cell4thBottomDryer1Current.AddParagraph($"{fourthBottomDryer1Age}");

            var cell4thBottomDryer1Goal = row4thBottomDryer.Cells[3];
            cell4thBottomDryer1Goal.Format.Font.Size = 6.5;
            cell4thBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 13);
            cell4thBottomDryer1Goal.AddParagraph($"{fourthBottomDryerlGoal.Goal1}");

            var cell4thBottomDryer1Info = row4thBottomDryer2.Cells[1];
            cell4thBottomDryer1Info.Format.Font.Size = 6.5;
            cell4thBottomDryer1Info.MergeRight = 2;
            cell4thBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer1Info.Shading.Color = Colors.LightBlue;
            if (clothing4thBottomDryer1 != null)
            {
                cell4thBottomDryer1Info.AddParagraph($"{clothing4thBottomDryer1.Manufacturer} {clothing4thBottomDryer1.Serial_Number}");
            }
            else cell4thBottomDryer1Info.AddParagraph("NA");
            #endregion

            #region 4th Bottom PM 2
            // ----- 4th Bottom Position PM 2 ----- //
            var cell4thBottomDryer2Past = row4thBottomDryer.Cells[4];
            cell4thBottomDryer2Past.Format.Font.Size = 6.5;
            cell4thBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer2Past.AddParagraph($"100");

            var cell4thBottomDryer2Current = row4thBottomDryer.Cells[5];
            cell4thBottomDryer2Current.Format.Font.Size = 6.5;
            cell4thBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryer2Age = 0;
            if (clothing4thBottomDryer2 != null)
            {
                fourthBottomDryer2Age = clothing4thBottomDryer2.Age != null ? Convert.ToInt32(clothing4thBottomDryer2.Age) : 0;
            }

            cell4thBottomDryer2Current.AddParagraph($"{fourthBottomDryer2Age}");

            var cell4thBottomDryer2Goal = row4thBottomDryer.Cells[6];
            cell4thBottomDryer2Goal.Format.Font.Size = 6.5;
            cell4thBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 13);
            cell4thBottomDryer2Goal.AddParagraph($"{fourthBottomDryer2Goal.Goal1}");

            var cell4thBottomDryer2Info = row4thBottomDryer2.Cells[4];
            cell4thBottomDryer2Info.Format.Font.Size = 6.5;
            cell4thBottomDryer2Info.MergeRight = 2;
            cell4thBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer2Info.Shading.Color = Colors.LightBlue;
            if (clothing4thBottomDryer2 != null)
            {
                cell4thBottomDryer2Info.AddParagraph($"{clothing4thBottomDryer2.Manufacturer} {clothing4thBottomDryer2.Serial_Number}");
            }
            else cell4thBottomDryer2Info.AddParagraph("NA");
            #endregion

            #region 4th Bottom PM 3
            // ----- 4th Bottom Position PM 3 ----- //
            var cell4thBottomDryer3Past = row4thBottomDryer.Cells[7];
            cell4thBottomDryer3Past.Format.Font.Size = 6.5;
            cell4thBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer3Past.AddParagraph($"100");

            var cell4thBottomDryer3Current = row4thBottomDryer.Cells[8];
            cell4thBottomDryer3Current.Format.Font.Size = 6.5;
            cell4thBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryer3Age = 0;
            if (clothing4thBottomDryer3 != null)
            {
                fourthBottomDryer3Age = clothing4thBottomDryer3.Age != null ? Convert.ToInt32(clothing4thBottomDryer3.Age) : 0;
            }

            cell4thBottomDryer3Current.AddParagraph($"{fourthBottomDryer3Age}");

            var cell4thBottomDryer3Goal = row4thBottomDryer.Cells[9];
            cell4thBottomDryer3Goal.Format.Font.Size = 6.5;
            cell4thBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryer3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 13);
            cell4thBottomDryer3Goal.AddParagraph($"{fourthBottomDryer3Goal.Goal1}");

            var cell4thBottomDryer3Info = row4thBottomDryer2.Cells[7];
            cell4thBottomDryer3Info.Format.Font.Size = 6.5;
            cell4thBottomDryer3Info.MergeRight = 2;
            cell4thBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer3Info.Shading.Color = Colors.LightBlue;
            if (clothing4thBottomDryer3 != null)
            {
                cell4thBottomDryer3Info.AddParagraph($"{clothing4thBottomDryer3.Manufacturer} {clothing4thBottomDryer3.Serial_Number}");
            }
            else cell4thBottomDryer3Info.AddParagraph("NA");
            #endregion

            #region 4th Bottom PM 4
            // ----- 4th Bottom Position PM 4 ----- //
            var cell4thBottomDryer4Past = row4thBottomDryer.Cells[10];
            cell4thBottomDryer4Past.Format.Font.Size = 6.5;
            cell4thBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer4Past.AddParagraph($"100");

            var cell4thBottomDryer4Current = row4thBottomDryer.Cells[11];
            cell4thBottomDryer4Current.Format.Font.Size = 6.5;
            cell4thBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryer4Age = 0;
            if (clothing4thBottomDryer4 != null)
            {
                fourthBottomDryer4Age = clothing4thBottomDryer4.Age != null ? Convert.ToInt32(clothing4thBottomDryer4.Age) : 0;
            }

            cell4thBottomDryer4Current.AddParagraph($"{fourthBottomDryer4Age}");

            var cell4thBottomDryer4Goal = row4thBottomDryer.Cells[12];
            cell4thBottomDryer4Goal.Format.Font.Size = 6.5;
            cell4thBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            var fourthBottomDryer4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 13);
            cell4thBottomDryer4Goal.AddParagraph($"{fourthBottomDryer4Goal.Goal1}");

            var cell4thBottomDryer4Info = row4thBottomDryer2.Cells[10];
            cell4thBottomDryer4Info.Format.Font.Size = 6.5;
            cell4thBottomDryer4Info.MergeRight = 2;
            cell4thBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer4Info.Shading.Color = Colors.LightBlue;
            if (clothing4thBottomDryer4 != null)
            {
                cell4thBottomDryer4Info.AddParagraph($"{clothing4thBottomDryer4.Manufacturer} {clothing4thBottomDryer4.Serial_Number}");
            }
            else cell4thBottomDryer4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region Breast Roll Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingBreast1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 1);
            var clothingBreast2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 2);
            var clothingBreast3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 3);
            var clothingBreast4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 4);
            var rowBreast = weeklyPMTable.AddRow();
            rowBreast.Height = 10;
            var rowBreast2 = weeklyPMTable.AddRow();
            rowBreast2.Height = 10;

            //define Breast position row cells

            var cellBreast = rowBreast.Cells[0];
            cellBreast.Format.Font.Size = 6.5;
            cellBreast.VerticalAlignment = VerticalAlignment.Center;
            cellBreast.MergeDown = 1;
            cellBreast.AddParagraph("BREAST ROLL");

            #region Breast PM 1
            // ----- Breast Position PM 1 ----- //
            var cellBreast1Past = rowBreast.Cells[1];
            cellBreast1Past.Format.Font.Size = 6.5;
            cellBreast1Past.VerticalAlignment = VerticalAlignment.Center;
            cellBreast1Past.AddParagraph($"100");

            var cellBreast1Current = rowBreast.Cells[2];
            cellBreast1Current.Format.Font.Size = 6.5;
            cellBreast1Current.VerticalAlignment = VerticalAlignment.Center;
            var breast1Age = 0;
            if (clothingBreast1 != null)
            {
                breast1Age = clothingBreast1.Age != null ? Convert.ToInt32(clothingBreast1.Age) : 0;
            }

            cellBreast1Current.AddParagraph($"{breast1Age}");

            var cellBreast1Goal = rowBreast.Cells[3];
            cellBreast1Goal.Format.Font.Size = 6.5;
            cellBreast1Goal.VerticalAlignment = VerticalAlignment.Center;
            var breastlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 14);
            cellBreast1Goal.AddParagraph($"{breastlGoal.Goal1}");

            var cellBreast1Info = rowBreast2.Cells[1];
            cellBreast1Info.Format.Font.Size = 6.5;
            cellBreast1Info.MergeRight = 2;
            cellBreast1Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast1Info.Shading.Color = Colors.LightBlue;
            if (clothingBreast1 != null)
            {
                cellBreast1Info.AddParagraph($"{clothingBreast1.Manufacturer} {clothingBreast1.Serial_Number}");
            }
            else cellBreast1Info.AddParagraph("NA");
            #endregion

            #region Breast PM 2
            // ----- Breast Position PM 2 ----- //
            var cellBreast2Past = rowBreast.Cells[4];
            cellBreast2Past.Format.Font.Size = 6.5;
            cellBreast2Past.VerticalAlignment = VerticalAlignment.Center;
            cellBreast2Past.AddParagraph($"100");

            var cellBreast2Current = rowBreast.Cells[5];
            cellBreast2Current.Format.Font.Size = 6.5;
            cellBreast2Current.VerticalAlignment = VerticalAlignment.Center;
            var breast2Age = 0;
            if (clothingBreast2 != null)
            {
                breast2Age = clothingBreast2.Age != null ? Convert.ToInt32(clothingBreast2.Age) : 0;
            }

            cellBreast2Current.AddParagraph($"{breast2Age}");

            var cellBreast2Goal = rowBreast.Cells[6];
            cellBreast2Goal.Format.Font.Size = 6.5;
            cellBreast2Goal.VerticalAlignment = VerticalAlignment.Center;
            var breast2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 14);
            cellBreast2Goal.AddParagraph($"{breast2Goal.Goal1}");

            var cellBreast2Info = rowBreast2.Cells[4];
            cellBreast2Info.Format.Font.Size = 6.5;
            cellBreast2Info.MergeRight = 2;
            cellBreast2Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast2Info.Shading.Color = Colors.LightBlue;
            if (clothingBreast2 != null)
            {
                cellBreast2Info.AddParagraph($"{clothingBreast2.Manufacturer} {clothingBreast2.Serial_Number}");
            }
            else cellBreast2Info.AddParagraph("NA");
            #endregion

            #region Breast PM 3
            // ----- Breast Position PM 3 ----- //
            var cellBreast3Past = rowBreast.Cells[7];
            cellBreast3Past.Format.Font.Size = 6.5;
            cellBreast3Past.VerticalAlignment = VerticalAlignment.Center;
            cellBreast3Past.AddParagraph($"100");

            var cellBreast3Current = rowBreast.Cells[8];
            cellBreast3Current.Format.Font.Size = 6.5;
            cellBreast3Current.VerticalAlignment = VerticalAlignment.Center;
            var breast3Age = 0;
            if (clothingBreast3 != null)
            {
                breast3Age = clothingBreast3.Age != null ? Convert.ToInt32(clothingBreast3.Age) : 0;
            }

            cellBreast3Current.AddParagraph($"{breast3Age}");

            var cellBreast3Goal = rowBreast.Cells[9];
            cellBreast3Goal.Format.Font.Size = 6.5;
            cellBreast3Goal.VerticalAlignment = VerticalAlignment.Center;
            var breast3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 14);
            cellBreast3Goal.AddParagraph($"{breast3Goal.Goal1}");

            var cellBreast3Info = rowBreast2.Cells[7];
            cellBreast3Info.Format.Font.Size = 6.5;
            cellBreast3Info.MergeRight = 2;
            cellBreast3Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast3Info.Shading.Color = Colors.LightBlue;
            if (clothingBreast3 != null)
            {
                cellBreast3Info.AddParagraph($"{clothingBreast3.Manufacturer} {clothingBreast3.Serial_Number}");
            }
            else cellBreast3Info.AddParagraph("NA");
            #endregion

            #region Breast PM 4
            // ----- Breast Position PM 4 ----- //
            var cellBreast4Past = rowBreast.Cells[10];
            cellBreast4Past.Format.Font.Size = 6.5;
            cellBreast4Past.VerticalAlignment = VerticalAlignment.Center;
            cellBreast4Past.AddParagraph($"100");

            var cellBreast4Current = rowBreast.Cells[11];
            cellBreast4Current.Format.Font.Size = 6.5;
            cellBreast4Current.VerticalAlignment = VerticalAlignment.Center;
            var breast4Age = 0;
            if (clothingBreast4 != null)
            {
                breast4Age = clothingBreast4.Age != null ? Convert.ToInt32(clothingBreast4.Age) : 0;
            }

            cellBreast4Current.AddParagraph($"{breast4Age}");

            var cellBreast4Goal = rowBreast.Cells[12];
            cellBreast4Goal.Format.Font.Size = 6.5;
            cellBreast4Goal.VerticalAlignment = VerticalAlignment.Center;
            var breast4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 14);
            cellBreast4Goal.AddParagraph($"{breast4Goal.Goal1}");

            var cellBreast4Info = rowBreast2.Cells[10];
            cellBreast4Info.Format.Font.Size = 6.5;
            cellBreast4Info.MergeRight = 2;
            cellBreast4Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast4Info.Shading.Color = Colors.LightBlue;
            if (clothingBreast4 != null)
            {
                cellBreast4Info.AddParagraph($"{clothingBreast4.Manufacturer} {clothingBreast4.Serial_Number}");
            }
            else cellBreast4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region Dandy Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingDandy1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy" && c.PM_Number == 1);
            var clothingDandy2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy" && c.PM_Number == 2);
            var clothingDandy3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy" && c.PM_Number == 3);
            var clothingDandy4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy" && c.PM_Number == 4);
            var rowDandy = weeklyPMTable.AddRow();
            rowDandy.Height = 10;
            var rowDandy2 = weeklyPMTable.AddRow();
            rowDandy2.Height = 10;

            //define Dandy position row cells

            var cellDandy = rowDandy.Cells[0];
            cellDandy.Format.Font.Size = 6.5;
            cellDandy.VerticalAlignment = VerticalAlignment.Center;
            cellDandy.MergeDown = 1;
            cellDandy.AddParagraph("DANDY");

            #region Dandy PM 1
            // ----- Dandy Position PM 1 ----- //
            var cellDandy1Past = rowDandy.Cells[1];
            cellDandy1Past.Format.Font.Size = 6.5;
            cellDandy1Past.VerticalAlignment = VerticalAlignment.Center;
            cellDandy1Past.AddParagraph($"100");

            var cellDandy1Current = rowDandy.Cells[2];
            cellDandy1Current.Format.Font.Size = 6.5;
            cellDandy1Current.VerticalAlignment = VerticalAlignment.Center;
            var dandy1Age = 0;
            if (clothingDandy1 != null)
            {
                dandy1Age = clothingDandy1.Age != null ? Convert.ToInt32(clothingDandy1.Age) : 0;
            }

            cellDandy1Current.AddParagraph($"{dandy1Age}");

            var cellDandy1Goal = rowDandy.Cells[3];
            cellDandy1Goal.Format.Font.Size = 6.5;
            cellDandy1Goal.VerticalAlignment = VerticalAlignment.Center;
            var dandylGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 15);
            cellDandy1Goal.AddParagraph($"{dandylGoal.Goal1}");

            var cellDandy1Info = rowDandy2.Cells[1];
            cellDandy1Info.Format.Font.Size = 6.5;
            cellDandy1Info.MergeRight = 2;
            cellDandy1Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy1Info.Shading.Color = Colors.LightBlue;
            if (clothingDandy1 != null)
            {
                cellDandy1Info.AddParagraph($"{clothingDandy1.Manufacturer} {clothingDandy1.Serial_Number}");
            }
            else cellDandy1Info.AddParagraph("NA");
            #endregion

            #region Dandy PM 2
            // ----- Dandy Position PM 2 ----- //
            var cellDandy2Past = rowDandy.Cells[4];
            cellDandy2Past.Format.Font.Size = 6.5;
            cellDandy2Past.VerticalAlignment = VerticalAlignment.Center;
            cellDandy2Past.AddParagraph($"100");

            var cellDandy2Current = rowDandy.Cells[5];
            cellDandy2Current.Format.Font.Size = 6.5;
            cellDandy2Current.VerticalAlignment = VerticalAlignment.Center;
            var dandy2Age = 0;
            if (clothingDandy2 != null)
            {
                dandy2Age = clothingDandy2.Age != null ? Convert.ToInt32(clothingDandy2.Age) : 0;
            }

            cellDandy2Current.AddParagraph($"{dandy2Age}");

            var cellDandy2Goal = rowDandy.Cells[6];
            cellDandy2Goal.Format.Font.Size = 6.5;
            cellDandy2Goal.VerticalAlignment = VerticalAlignment.Center;
            var dandy2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 15);
            cellDandy2Goal.AddParagraph($"{dandy2Goal.Goal1}");

            var cellDandy2Info = rowDandy2.Cells[4];
            cellDandy2Info.Format.Font.Size = 6.5;
            cellDandy2Info.MergeRight = 2;
            cellDandy2Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy2Info.Shading.Color = Colors.LightBlue;
            if (clothingDandy2 != null)
            {
                cellDandy2Info.AddParagraph($"{clothingDandy2.Manufacturer} {clothingDandy2.Serial_Number}");
            }
            else cellDandy2Info.AddParagraph("NA");
            #endregion

            #region Dandy PM 3
            // ----- Dandy Position PM 3 ----- //
            var cellDandy3Past = rowDandy.Cells[7];
            cellDandy3Past.Format.Font.Size = 6.5;
            cellDandy3Past.VerticalAlignment = VerticalAlignment.Center;
            cellDandy3Past.AddParagraph($"100");

            var cellDandy3Current = rowDandy.Cells[8];
            cellDandy3Current.Format.Font.Size = 6.5;
            cellDandy3Current.VerticalAlignment = VerticalAlignment.Center;
            var dandy3Age = 0;
            if (clothingDandy3 != null)
            {
                dandy3Age = clothingDandy3.Age != null ? Convert.ToInt32(clothingDandy3.Age) : 0;
            }

            cellDandy3Current.AddParagraph($"{dandy3Age}");

            var cellDandy3Goal = rowDandy.Cells[9];
            cellDandy3Goal.Format.Font.Size = 6.5;
            cellDandy3Goal.VerticalAlignment = VerticalAlignment.Center;
            var dandy3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 15);
            cellDandy3Goal.AddParagraph($"{dandy3Goal.Goal1}");

            var cellDandy3Info = rowDandy2.Cells[7];
            cellDandy3Info.Format.Font.Size = 6.5;
            cellDandy3Info.MergeRight = 2;
            cellDandy3Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy3Info.Shading.Color = Colors.LightBlue;
            if (clothingDandy3 != null)
            {
                cellDandy3Info.AddParagraph($"{clothingDandy3.Manufacturer} {clothingDandy3.Serial_Number}");
            }
            else cellDandy3Info.AddParagraph("NA");
            #endregion

            #region Dandy PM 4
            // ----- Dandy Position PM 4 ----- //
            var cellDandy4Past = rowDandy.Cells[10];
            cellDandy4Past.Format.Font.Size = 6.5;
            cellDandy4Past.VerticalAlignment = VerticalAlignment.Center;
            cellDandy4Past.AddParagraph($"100");

            var cellDandy4Current = rowDandy.Cells[11];
            cellDandy4Current.Format.Font.Size = 6.5;
            cellDandy4Current.VerticalAlignment = VerticalAlignment.Center;
            var dandy4Age = 0;
            if (clothingDandy4 != null)
            {
                dandy4Age = clothingDandy4.Age != null ? Convert.ToInt32(clothingDandy4.Age) : 0;
            }

            cellDandy4Current.AddParagraph($"{dandy4Age}");

            var cellDandy4Goal = rowDandy.Cells[12];
            cellDandy4Goal.Format.Font.Size = 6.5;
            cellDandy4Goal.VerticalAlignment = VerticalAlignment.Center;
            var dandy4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 15);
            cellDandy4Goal.AddParagraph($"{dandy4Goal.Goal1}");

            var cellDandy4Info = rowDandy2.Cells[10];
            cellDandy4Info.Format.Font.Size = 6.5;
            cellDandy4Info.MergeRight = 2;
            cellDandy4Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy4Info.Shading.Color = Colors.LightBlue;
            if (clothingDandy4 != null)
            {
                cellDandy4Info.AddParagraph($"{clothingDandy4.Manufacturer} {clothingDandy4.Serial_Number}");
            }
            else cellDandy4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region Lumpbreaker Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingLumpbreaker1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker" && c.PM_Number == 1);
            var clothingLumpbreaker2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker" && c.PM_Number == 2);
            var clothingLumpbreaker3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker" && c.PM_Number == 3);
            var clothingLumpbreaker4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker" && c.PM_Number == 4);
            var rowLumpbreaker = weeklyPMTable.AddRow();
            rowLumpbreaker.Height = 10;
            var rowLumpbreaker2 = weeklyPMTable.AddRow();
            rowLumpbreaker2.Height = 10;

            //define Dandy position row cells

            var cellLumpbreaker = rowLumpbreaker.Cells[0];
            cellLumpbreaker.Format.Font.Size = 6.5;
            cellLumpbreaker.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker.MergeDown = 1;
            cellLumpbreaker.AddParagraph("LUMPBREAKER");

            #region Lumpbreaker PM 1
            // ----- Lumpbreaker Position PM 1 ----- //
            var cellLumpbreaker1Past = rowLumpbreaker.Cells[1];
            cellLumpbreaker1Past.Format.Font.Size = 6.5;
            cellLumpbreaker1Past.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker1Past.AddParagraph($"100");

            var cellLumpbreaker1Current = rowLumpbreaker.Cells[2];
            cellLumpbreaker1Current.Format.Font.Size = 6.5;
            cellLumpbreaker1Current.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreaker1Age = 0;
            if (clothingLumpbreaker1 != null)
            {
                lumpbreaker1Age = clothingLumpbreaker1.Age != null ? Convert.ToInt32(clothingLumpbreaker1.Age) : 0;
            }

            cellLumpbreaker1Current.AddParagraph($"{lumpbreaker1Age}");

            var cellLumpbreaker1Goal = rowLumpbreaker.Cells[3];
            cellLumpbreaker1Goal.Format.Font.Size = 6.5;
            cellLumpbreaker1Goal.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreakerlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 16);
            cellLumpbreaker1Goal.AddParagraph($"{lumpbreakerlGoal.Goal1}");

            var cellLumpbreaker1Info = rowLumpbreaker2.Cells[1];
            cellLumpbreaker1Info.Format.Font.Size = 6.5;
            cellLumpbreaker1Info.MergeRight = 2;
            cellLumpbreaker1Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker1Info.Shading.Color = Colors.LightBlue;
            if (clothingLumpbreaker1 != null)
            {
                cellLumpbreaker1Info.AddParagraph($"{clothingLumpbreaker1.Manufacturer} {clothingLumpbreaker1.Serial_Number}");
            }
            else cellLumpbreaker1Info.AddParagraph("NA");
            #endregion

            #region Lumpbreaker PM 2
            // ----- Lumpbreaker Position PM 2 ----- //
            var cellLumpbreaker2Past = rowLumpbreaker.Cells[4];
            cellLumpbreaker2Past.Format.Font.Size = 6.5;
            cellLumpbreaker2Past.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker2Past.AddParagraph($"100");

            var cellLumpbreaker2Current = rowLumpbreaker.Cells[5];
            cellLumpbreaker2Current.Format.Font.Size = 6.5;
            cellLumpbreaker2Current.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreaker2Age = 0;
            if (clothingLumpbreaker2 != null)
            {
                lumpbreaker2Age = clothingLumpbreaker2.Age != null ? Convert.ToInt32(clothingLumpbreaker2.Age) : 0;
            }

            cellLumpbreaker2Current.AddParagraph($"{lumpbreaker2Age}");

            var cellLumpbreaker2Goal = rowLumpbreaker.Cells[6];
            cellLumpbreaker2Goal.Format.Font.Size = 6.5;
            cellLumpbreaker2Goal.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreaker2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 16);
            cellLumpbreaker2Goal.AddParagraph($"{lumpbreaker2Goal.Goal1}");

            var cellLumpbreaker2Info = rowLumpbreaker2.Cells[4];
            cellLumpbreaker2Info.Format.Font.Size = 6.5;
            cellLumpbreaker2Info.MergeRight = 2;
            cellLumpbreaker2Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker2Info.Shading.Color = Colors.LightBlue;
            if (clothingLumpbreaker2 != null)
            {
                cellLumpbreaker2Info.AddParagraph($"{clothingLumpbreaker2.Manufacturer} {clothingLumpbreaker2.Serial_Number}");
            }
            else cellLumpbreaker2Info.AddParagraph("NA");
            #endregion

            #region Lumpbreaker PM 3
            // ----- Lumpbreaker Position PM 3 ----- //
            var cellLumpbreaker3Past = rowLumpbreaker.Cells[7];
            cellLumpbreaker3Past.Format.Font.Size = 6.5;
            cellLumpbreaker3Past.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker3Past.AddParagraph($"100");

            var cellLumpbreaker3Current = rowLumpbreaker.Cells[8];
            cellLumpbreaker3Current.Format.Font.Size = 6.5;
            cellLumpbreaker3Current.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreaker3Age = 0;
            if (clothingLumpbreaker3 != null)
            {
                lumpbreaker3Age = clothingLumpbreaker3.Age != null ? Convert.ToInt32(clothingLumpbreaker3.Age) : 0;
            }

            cellLumpbreaker3Current.AddParagraph($"{lumpbreaker3Age}");

            var cellLumpbreaker3Goal = rowLumpbreaker.Cells[9];
            cellLumpbreaker3Goal.Format.Font.Size = 6.5;
            cellLumpbreaker3Goal.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreaker3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 16);
            cellLumpbreaker3Goal.AddParagraph($"{lumpbreaker3Goal.Goal1}");

            var cellLumpbreaker3Info = rowLumpbreaker2.Cells[7];
            cellLumpbreaker3Info.Format.Font.Size = 6.5;
            cellLumpbreaker3Info.MergeRight = 2;
            cellLumpbreaker3Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker3Info.Shading.Color = Colors.LightBlue;
            if (clothingLumpbreaker3 != null)
            {
                cellLumpbreaker3Info.AddParagraph($"{clothingLumpbreaker3.Manufacturer} {clothingLumpbreaker3.Serial_Number}");
            }
            else cellLumpbreaker3Info.AddParagraph("NA");
            #endregion

            #region Lumpbreaker PM 4
            // ----- Lumpbreaker Position PM 4 ----- //
            var cellLumpbreaker4Past = rowLumpbreaker.Cells[10];
            cellLumpbreaker4Past.Format.Font.Size = 6.5;
            cellLumpbreaker4Past.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker4Past.AddParagraph($"100");

            var cellLumpbreaker4Current = rowLumpbreaker.Cells[11];
            cellLumpbreaker4Current.Format.Font.Size = 6.5;
            cellLumpbreaker4Current.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreaker4Age = 0;
            if (clothingLumpbreaker4 != null)
            {
                lumpbreaker4Age = clothingLumpbreaker4.Age != null ? Convert.ToInt32(clothingLumpbreaker4.Age) : 0;
            }

            cellLumpbreaker4Current.AddParagraph($"{lumpbreaker4Age}");

            var cellLumpbreaker4Goal = rowLumpbreaker.Cells[12];
            cellLumpbreaker4Goal.Format.Font.Size = 6.5;
            cellLumpbreaker4Goal.VerticalAlignment = VerticalAlignment.Center;
            var lumpbreaker4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 16);
            cellLumpbreaker4Goal.AddParagraph($"{lumpbreaker4Goal.Goal1}");

            var cellLumpbreaker4Info = rowLumpbreaker2.Cells[10];
            cellLumpbreaker4Info.Format.Font.Size = 6.5;
            cellLumpbreaker4Info.MergeRight = 2;
            cellLumpbreaker4Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker4Info.Shading.Color = Colors.LightBlue;
            if (clothingLumpbreaker4 != null)
            {
                cellLumpbreaker4Info.AddParagraph($"{clothingLumpbreaker4.Manufacturer} {clothingLumpbreaker4.Serial_Number}");
            }
            else cellLumpbreaker4Info.AddParagraph("NA");
            #endregion

            #endregion

            #region Suction Pickup Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSuction1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup" && c.PM_Number == 1);
            var clothingSuction2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup" && c.PM_Number == 2);
            var clothingSuction3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup" && c.PM_Number == 3);
            var clothingSuction4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup" && c.PM_Number == 4);
            var rowSuction = weeklyPMTable.AddRow();
            rowSuction.Height = 10;
            var rowSuction2 = weeklyPMTable.AddRow();
            rowSuction2.Height = 10;

            //define Dandy position row cells

            var cellSuction = rowSuction.Cells[0];
            cellSuction.Format.Font.Size = 6.5;
            cellSuction.VerticalAlignment = VerticalAlignment.Center;
            cellSuction.MergeDown = 1;
            cellSuction.AddParagraph("SUCTION PICKUP");

            #region Suction Pickup PM 1
            // ----- Suction Position PM 1 ----- //
            var cellSuction1Past = rowSuction.Cells[1];
            cellSuction1Past.Format.Font.Size = 6.5;
            cellSuction1Past.VerticalAlignment = VerticalAlignment.Center;
            cellSuction1Past.AddParagraph($"100");

            var cellSuction1Current = rowSuction.Cells[2];
            cellSuction1Current.Format.Font.Size = 6.5;
            cellSuction1Current.VerticalAlignment = VerticalAlignment.Center;
            var suction1Age = 0;
            if (clothingSuction1 != null)
            {
                suction1Age = clothingSuction1.Age != null ? Convert.ToInt32(clothingSuction1.Age) : 0;
            }

            cellSuction1Current.AddParagraph($"{suction1Age}");

            var cellSuction1Goal = rowSuction.Cells[3];
            cellSuction1Goal.Format.Font.Size = 6.5;
            cellSuction1Goal.VerticalAlignment = VerticalAlignment.Center;
            var suctionlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 17);
            cellSuction1Goal.AddParagraph($"{suctionlGoal.Goal1}");

            var cellSuction1Info = rowSuction2.Cells[1];
            cellSuction1Info.Format.Font.Size = 6.5;
            cellSuction1Info.MergeRight = 2;
            cellSuction1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction1Info.Shading.Color = Colors.LightBlue;
            if (clothingSuction1 != null)
            {
                cellSuction1Info.AddParagraph($"{clothingSuction1.Manufacturer} {clothingSuction1.Serial_Number}");
            }
            else cellSuction1Info.AddParagraph("NA");
            #endregion

            #region Suction Pickup PM 2
            // ----- Suction Position PM 2 ----- //
            var cellSuction2Past = rowSuction.Cells[4];
            cellSuction2Past.Format.Font.Size = 6.5;
            cellSuction2Past.VerticalAlignment = VerticalAlignment.Center;
            cellSuction2Past.AddParagraph($"100");

            var cellSuction2Current = rowSuction.Cells[5];
            cellSuction2Current.Format.Font.Size = 6.5;
            cellSuction2Current.VerticalAlignment = VerticalAlignment.Center;
            var suction2Age = 0;
            if (clothingSuction2 != null)
            {
                suction2Age = clothingSuction2.Age != null ? Convert.ToInt32(clothingSuction2.Age) : 0;
            }

            cellSuction2Current.AddParagraph($"{suction2Age}");

            var cellSuction2Goal = rowSuction.Cells[6];
            cellSuction2Goal.Format.Font.Size = 6.5;
            cellSuction2Goal.VerticalAlignment = VerticalAlignment.Center;
            var suction2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 17);
            cellSuction2Goal.AddParagraph($"{suction2Goal.Goal1}");

            var cellSuction2Info = rowSuction2.Cells[4];
            cellSuction2Info.Format.Font.Size = 6.5;
            cellSuction2Info.MergeRight = 2;
            cellSuction2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction2Info.Shading.Color = Colors.LightBlue;
            if (clothingSuction2 != null)
            {
                cellSuction2Info.AddParagraph($"{clothingSuction2.Manufacturer} {clothingSuction2.Serial_Number}");
            }
            else cellSuction2Info.AddParagraph("NA");
            #endregion

            #region  Suction Pickup PM 3
            // ----- Suction Position PM 3 ----- //
            var cellSuction3Past = rowSuction.Cells[7];
            cellSuction3Past.Format.Font.Size = 6.5;
            cellSuction3Past.VerticalAlignment = VerticalAlignment.Center;
            cellSuction3Past.AddParagraph($"100");

            var cellSuction3Current = rowSuction.Cells[8];
            cellSuction3Current.Format.Font.Size = 6.5;
            cellSuction3Current.VerticalAlignment = VerticalAlignment.Center;
            var suction3Age = 0;
            if (clothingSuction3 != null)
            {
                suction3Age = clothingSuction3.Age != null ? Convert.ToInt32(clothingSuction3.Age) : 0;
            }

            cellSuction3Current.AddParagraph($"{suction3Age}");

            var cellSuction3Goal = rowSuction.Cells[9];
            cellSuction3Goal.Format.Font.Size = 6.5;
            cellSuction3Goal.VerticalAlignment = VerticalAlignment.Center;
            var suction3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 17);
            cellSuction3Goal.AddParagraph($"{suction3Goal.Goal1}");

            var cellSuction3Info = rowSuction2.Cells[7];
            cellSuction3Info.Format.Font.Size = 6.5;
            cellSuction3Info.MergeRight = 2;
            cellSuction3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction3Info.Shading.Color = Colors.LightBlue;
            if (clothingSuction3 != null)
            {
                cellSuction3Info.AddParagraph($"{clothingSuction3.Manufacturer} {clothingSuction3.Serial_Number}");
            }
            else cellSuction3Info.AddParagraph("NA");
            #endregion

            #region Suction Pickup PM 4
            // ----- Suction Position PM 4 ----- //
            var cellSuction4Past = rowSuction.Cells[10];
            cellSuction4Past.Format.Font.Size = 6.5;
            cellSuction4Past.VerticalAlignment = VerticalAlignment.Center;
            cellSuction4Past.AddParagraph($"100");

            var cellSuction4Current = rowSuction.Cells[11];
            cellSuction4Current.Format.Font.Size = 6.5;
            cellSuction4Current.VerticalAlignment = VerticalAlignment.Center;
            var suction4Age = 0;
            if (clothingSuction4 != null)
            {
                suction4Age = clothingSuction4.Age != null ? Convert.ToInt32(clothingSuction4.Age) : 0;
            }

            cellSuction4Current.AddParagraph($"{suction4Age}");

            var cellSuction4Goal = rowSuction.Cells[12];
            cellSuction4Goal.Format.Font.Size = 6.5;
            cellSuction4Goal.VerticalAlignment = VerticalAlignment.Center;
            var suction4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 17);
            cellSuction4Goal.AddParagraph($"{suction4Goal.Goal1}");

            var cellSuction4Info = rowSuction2.Cells[10];
            cellSuction4Info.Format.Font.Size = 6.5;
            cellSuction4Info.MergeRight = 2;
            cellSuction4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction4Info.Shading.Color = Colors.LightBlue;
            if (clothingSuction4 != null)
            {
                cellSuction4Info.AddParagraph($"{clothingSuction4.Manufacturer} {clothingSuction4.Serial_Number}");
            }
            else cellSuction4Info.AddParagraph("NA");
            #endregion
            #endregion // Suction Pickup Position

            #region 1st Press Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stPressTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top" && c.PM_Number == 1);
            var clothing1stPressTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top" && c.PM_Number == 2);
            var clothing1stPressTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top" && c.PM_Number == 3);
            var clothing1stPressTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top" && c.PM_Number == 4);
            var row1stPressTop = weeklyPMTable.AddRow();
            row1stPressTop.Height = 10;
            var row1stPressTop2 = weeklyPMTable.AddRow();
            row1stPressTop2.Height = 10;

            //define Dandy position row cells

            var cell1stPressTop = row1stPressTop.Cells[0];
            cell1stPressTop.Format.Font.Size = 6.5;
            cell1stPressTop.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop.MergeDown = 1;
            cell1stPressTop.AddParagraph("1ST PRESS TOP");

            #region 1st Press Top PM 1
            // ----- 1st Press Top Position PM 1 ----- //
            var cell1stPressTop1Past = row1stPressTop.Cells[1];
            cell1stPressTop1Past.Format.Font.Size = 6.5;
            cell1stPressTop1Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop1Past.AddParagraph($"100");

            var cell1stPressTop1Current = row1stPressTop.Cells[2];
            cell1stPressTop1Current.Format.Font.Size = 6.5;
            cell1stPressTop1Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressTop1Age = 0;
            if (clothing1stPressTop1 != null)
            {
                firstPressTop1Age = clothing1stPressTop1.Age != null ? Convert.ToInt32(clothing1stPressTop1.Age) : 0;
            }
            cell1stPressTop1Current.AddParagraph($"{firstPressTop1Age}");

            var cell1stPressTop1Goal = row1stPressTop.Cells[3];
            cell1stPressTop1Goal.Format.Font.Size = 6.5;
            cell1stPressTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressToplGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 18);
            cell1stPressTop1Goal.AddParagraph($"{firstPressToplGoal.Goal1}");

            var cell1stPressTop1Info = row1stPressTop2.Cells[1];
            cell1stPressTop1Info.Format.Font.Size = 6.5;
            cell1stPressTop1Info.MergeRight = 2;
            cell1stPressTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop1Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressTop1 != null)
            {
                cell1stPressTop1Info.AddParagraph($"{clothing1stPressTop1.Manufacturer} {clothing1stPressTop1.Serial_Number}");
            }
            else cell1stPressTop1Info.AddParagraph("NA");
            #endregion

            #region 1st Press Top PM 2
            // ----- 1st Press Top Position PM 2 ----- //
            var cell1stPressTop2Past = row1stPressTop.Cells[4];
            cell1stPressTop2Past.Format.Font.Size = 6.5;
            cell1stPressTop2Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop2Past.AddParagraph($"100");

            var cell1stPressTop2Current = row1stPressTop.Cells[5];
            cell1stPressTop2Current.Format.Font.Size = 6.5;
            cell1stPressTop2Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressTop2Age = 0;
            if (clothing1stPressTop2 != null)
            {
                firstPressTop2Age = clothing1stPressTop2.Age != null ? Convert.ToInt32(clothing1stPressTop2.Age) : 0;
            }
            cell1stPressTop2Current.AddParagraph($"{firstPressTop2Age}");

            var cell1stPressTop2Goal = row1stPressTop.Cells[6];
            cell1stPressTop2Goal.Format.Font.Size = 6.5;
            cell1stPressTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressTop2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 18);
            cell1stPressTop2Goal.AddParagraph($"{firstPressTop2Goal.Goal1}");

            var cell1stPressTop2Info = row1stPressTop2.Cells[4];
            cell1stPressTop2Info.Format.Font.Size = 6.5;
            cell1stPressTop2Info.MergeRight = 2;
            cell1stPressTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop2Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressTop2 != null)
            {
                cell1stPressTop2Info.AddParagraph($"{clothing1stPressTop2.Manufacturer} {clothing1stPressTop2.Serial_Number}");
            }
            else cell1stPressTop2Info.AddParagraph("NA");
            #endregion

            #region 1st Press Top PM 3
            // ----- 1st Press Top Position PM 3 ----- //
            var cell1stPressTop3Past = row1stPressTop.Cells[7];
            cell1stPressTop3Past.Format.Font.Size = 6.5;
            cell1stPressTop3Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop3Past.AddParagraph($"100");

            var cell1stPressTop3Current = row1stPressTop.Cells[8];
            cell1stPressTop3Current.Format.Font.Size = 6.5;
            cell1stPressTop3Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressTop3Age = 0;
            if (clothing1stPressTop3 != null)
            {
                firstPressTop3Age = clothing1stPressTop3.Age != null ? Convert.ToInt32(clothing1stPressTop3.Age) : 0;
            }
            cell1stPressTop3Current.AddParagraph($"{firstPressTop3Age}");

            var cell1stPressTop3Goal = row1stPressTop.Cells[9];
            cell1stPressTop3Goal.Format.Font.Size = 6.5;
            cell1stPressTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressTop3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 18);
            cell1stPressTop3Goal.AddParagraph($"{firstPressTop3Goal.Goal1}");

            var cell1stPressTop3Info = row1stPressTop2.Cells[7];
            cell1stPressTop3Info.Format.Font.Size = 6.5;
            cell1stPressTop3Info.MergeRight = 2;
            cell1stPressTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop3Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressTop3 != null)
            {
                cell1stPressTop3Info.AddParagraph($"{clothing1stPressTop3.Manufacturer} {clothing1stPressTop3.Serial_Number}");
            }
            else cell1stPressTop3Info.AddParagraph("NA");
            #endregion

            #region 1st Press Top PM 4
            // ----- 1st Press Top Position PM 4 ----- //
            var cell1stPressTop4Past = row1stPressTop.Cells[10];
            cell1stPressTop4Past.Format.Font.Size = 6.5;
            cell1stPressTop4Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop4Past.AddParagraph($"100");

            var cell1stPressTop4Current = row1stPressTop.Cells[11];
            cell1stPressTop4Current.Format.Font.Size = 6.5;
            cell1stPressTop4Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressTop4Age = 0;
            if (clothing1stPressTop4 != null)
            {
                firstPressTop4Age = clothing1stPressTop4.Age != null ? Convert.ToInt32(clothing1stPressTop4.Age) : 0;
            }
            cell1stPressTop4Current.AddParagraph($"{firstPressTop4Age}");

            var cell1stPressTop4Goal = row1stPressTop.Cells[12];
            cell1stPressTop4Goal.Format.Font.Size = 6.5;
            cell1stPressTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressTop4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 18);
            cell1stPressTop4Goal.AddParagraph($"{firstPressTop4Goal.Goal1}");

            var cell1stPressTop4Info = row1stPressTop2.Cells[10];
            cell1stPressTop4Info.Format.Font.Size = 6.5;
            cell1stPressTop4Info.MergeRight = 2;
            cell1stPressTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop4Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressTop4 != null)
            {
                cell1stPressTop4Info.AddParagraph($"{clothing1stPressTop4.Manufacturer} {clothing1stPressTop4.Serial_Number}");
            }
            else cell1stPressTop4Info.AddParagraph("NA");
            #endregion

            #endregion // 1st Press Top Position

            #region 1st Press Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 15;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stPressBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom" && c.PM_Number == 1);
            var clothing1stPressBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom" && c.PM_Number == 2);
            var clothing1stPressBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom" && c.PM_Number == 3);
            var clothing1stPressBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom" && c.PM_Number == 4);
            var row1stPressBottom = weeklyPMTable.AddRow();
            row1stPressBottom.Height = 10;
            var row1stPressBottom2 = weeklyPMTable.AddRow();
            row1stPressBottom2.Height = 10;

            //define 1st Press Bottom position row cells

            var cell1stPressBottom = row1stPressBottom.Cells[0];
            cell1stPressBottom.Format.Font.Size = 6.5;
            cell1stPressBottom.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom.MergeDown = 1;
            cell1stPressBottom.AddParagraph("1ST PRESS BOTTOM");

            #region 1st Press Bottom PM 1
            // ----- 1st Press Bottom Position PM 1 ----- //
            var cell1stPressBottom1Past = row1stPressBottom.Cells[1];
            cell1stPressBottom1Past.Format.Font.Size = 6.5;
            cell1stPressBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom1Past.AddParagraph($"100");

            var cell1stPressBottom1Current = row1stPressBottom.Cells[2];
            cell1stPressBottom1Current.Format.Font.Size = 6.5;
            cell1stPressBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom1Age = 0;
            if (clothing1stPressBottom1 != null)
            {
                firstPressBottom1Age = clothing1stPressBottom1.Age != null ? Convert.ToInt32(clothing1stPressBottom1.Age) : 0;
            }
            cell1stPressBottom1Current.AddParagraph($"{firstPressBottom1Age}");

            var cell1stPressBottom1Goal = row1stPressBottom.Cells[3];
            cell1stPressBottom1Goal.Format.Font.Size = 6.5;
            cell1stPressBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottomlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 19);
            cell1stPressBottom1Goal.AddParagraph($"{firstPressBottomlGoal.Goal1}");

            var cell1stPressBottom1Info = row1stPressBottom2.Cells[1];
            cell1stPressBottom1Info.Format.Font.Size = 6.5;
            cell1stPressBottom1Info.MergeRight = 2;
            cell1stPressBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom1Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressBottom1 != null)
            {
                cell1stPressBottom1Info.AddParagraph($"{clothing1stPressBottom1.Manufacturer} {clothing1stPressBottom1.Serial_Number}");
            }
            else cell1stPressBottom1Info.AddParagraph("NA");
            #endregion
            #endregion

            #region 1st Press Bottom PM 2

            #endregion

            #region 1st Press Bottom PM 3

            #endregion

            #region 1st Press Bottom PM 4

            #endregion

#endregion // 1st Press Bottom Position

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