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
            weeklyPMTable.Rows.Alignment = RowAlignment.Left;
            weeklyPMTable.Rows.LeftIndent = 20;

            #region Header Columns
            //define header columns
            var posCol = weeklyPMTable.AddColumn(Unit.FromCentimeter(3.5));
            posCol.Format.Alignment = ParagraphAlignment.Center;
            posCol.Format.Font.Size = 6.5;

            var pm1Col = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm1Col.Format.Alignment = ParagraphAlignment.Center;
            pm1Col.Format.Font.Size = 6.5;

            var pm1Colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm1Colb.Format.Alignment = ParagraphAlignment.Center;
            pm1Colb.Format.Font.Size = 6.5;

            var pm1Colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm1Colc.Format.Alignment = ParagraphAlignment.Center;
            pm1Colc.Format.Font.Size = 6.5;

            var pm2col = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm2col.Format.Alignment = ParagraphAlignment.Center;
            pm2col.Format.Font.Size = 6.5;

            var pm2colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm2colb.Format.Alignment = ParagraphAlignment.Center;
            pm2colb.Format.Font.Size = 6.5;

            var pm2colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm2colc.Format.Alignment = ParagraphAlignment.Center;
            pm2colc.Format.Font.Size = 6.5;

            var pm3col = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm3col.Format.Alignment = ParagraphAlignment.Center;
            pm3col.Format.Font.Size = 6.5;

            var pm3colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm3colb.Format.Alignment = ParagraphAlignment.Center;
            pm3colb.Format.Font.Size = 6.5;

            var pm3colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm3colc.Format.Alignment = ParagraphAlignment.Center;
            pm3colc.Format.Font.Size = 6.5;

            var pm4col = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm4col.Format.Alignment = ParagraphAlignment.Center;
            pm4col.Format.Font.Size = 6.5;

            var pm4colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
            pm4colb.Format.Alignment = ParagraphAlignment.Center;
            pm4colb.Format.Font.Size = 6.5;

            var pm4colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(1.5));
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
            pm1Header.AddParagraph("#1 PM Days Run");

            var pm2Header = headers.Cells[4];
            pm2Header.Format.Font.Size = 6.5;
            pm2Header.MergeRight = 2;
            pm2Header.VerticalAlignment = VerticalAlignment.Center;
            pm2Header.AddParagraph("#2 PM Days Run");

            var pm3Header = headers.Cells[7];
            pm3Header.Format.Font.Size = 6.5;
            pm3Header.MergeRight = 2;
            pm3Header.VerticalAlignment = VerticalAlignment.Center;
            pm3Header.AddParagraph("#3 PM Days Run");

            var pm4Header = headers.Cells[10];
            pm4Header.Format.Font.Size = 6.5;
            pm4Header.MergeRight = 2;
            pm4Header.VerticalAlignment = VerticalAlignment.Center;
            pm4Header.AddParagraph("#4 PM Days Run");

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
            var wire1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 2 && c.StatusID == 3).Select(c => c.Age).Average();
            wire1Past = wire1Past != null ? Math.Round((double)wire1Past) : 0;
            cellWire1Past.AddParagraph($"{wire1Past}");

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
            var wire2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 2 && c.StatusID == 3).Select(c => c.Age).Average();
            wire2Past = wire2Past != null ? Math.Round((double)wire2Past) : 0;
            cellWire2Past.AddParagraph($"{wire2Past}");

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
            var wire3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 2 && c.StatusID == 3).Select(c => c.Age).Average();
            wire3Past = wire3Past != null ? Math.Round((double)wire3Past) : 0;
            cellWire3Past.AddParagraph($"{wire3Past}");

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
            var wire4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 2 && c.StatusID == 3).Select(c => c.Age).Average();
            wire4Past = wire4Past != null ? Math.Round((double)wire4Past) : 0;
            cellWire4Past.AddParagraph($"{wire4Past}");

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
            spacer.Height = 12;
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
            var firstPress1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 3 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPress1Past = firstPress1Past != null ? Math.Round((double)firstPress1Past) : 0;
            cell1Press1Past.AddParagraph($"{firstPress1Past}");

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
            var firstPress2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 3 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPress2Past = firstPress2Past != null ? Math.Round((double)firstPress2Past) : 0;
            cell1Press2Past.AddParagraph($"{firstPress2Past}");

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
            var firstPress3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 3 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPress3Past = firstPress3Past != null ? Math.Round((double)firstPress3Past) : 0;
            cell1Press3Past.AddParagraph($"{firstPress3Past}");

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
            var firstPress4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 3 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPress4Past = firstPress4Past != null ? Math.Round((double)firstPress4Past) : 0;
            cell1Press4Past.AddParagraph($"{firstPress4Past}");

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
            spacer.Height = 12;
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
            var secondPress1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 4 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPress1Past = secondPress1Past != null ? Math.Round((double)secondPress1Past) : 0;
            cell2Press1Past.AddParagraph($"{secondPress1Past}");

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
            var secondPress2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 4 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPress2Past = secondPress2Past != null ? Math.Round((double)secondPress2Past) : 0;
            cell2Press2Past.AddParagraph($"{secondPress2Past}");

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
            var secondPress3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 4 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPress3Past = secondPress3Past != null ? Math.Round((double)secondPress3Past) : 0;
            cell2Press3Past.AddParagraph($"{secondPress3Past}");

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
            var secondPress4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 4 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPress4Past = secondPress4Past != null ? Math.Round((double)secondPress4Past) : 0;
            cell2Press4Past.AddParagraph($"{secondPress4Past}");

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
            spacer.Height = 12;
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
            var thirdPress1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 5 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPress1Past = thirdPress1Past != null ? Math.Round((double)thirdPress1Past) : 0;
            cell3Press1Past.AddParagraph($"{thirdPress1Past}");

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
            var thirdPress2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 5 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPress2Past = thirdPress2Past != null ? Math.Round((double)thirdPress2Past) : 0;
            cell3Press2Past.AddParagraph($"{thirdPress2Past}");

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
            var thirdPress3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 5 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPress3Past = thirdPress3Past != null ? Math.Round((double)thirdPress3Past) : 0;
            cell3Press3Past.AddParagraph($"{thirdPress3Past}");

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
            var thirdPress4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 5 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPress4Past = thirdPress4Past != null ? Math.Round((double)thirdPress4Past) : 0;
            cell3Press4Past.AddParagraph($"{thirdPress4Past}");

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

            #endregion // 3rd Press Position

            #region 1st Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
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
            var firstTopDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 6 && c.StatusID == 3).Select(c => c.Age).Average();
            firstTopDryer1Past = firstTopDryer1Past != null ? Math.Round((double)firstTopDryer1Past) : 0;
            cellFirstTopDryer1Past.AddParagraph($"{firstTopDryer1Past}");

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
            var firstTopDryer2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 6 && c.StatusID == 3).Select(c => c.Age).Average();
            firstTopDryer2Past = firstTopDryer2Past != null ? Math.Round((double)firstTopDryer2Past) : 0;
            cellFirstTopDryer2Past.AddParagraph($"{firstTopDryer2Past}");

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
            var firstTopDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 6 && c.StatusID == 3).Select(c => c.Age).Average();
            firstTopDryer3Past = firstTopDryer3Past != null ? Math.Round((double)firstTopDryer3Past) : 0;
            cellFirstTopDryer3Past.AddParagraph($"{firstTopDryer3Past}");

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
            var firstTopDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 6 && c.StatusID == 3).Select(c => c.Age).Average();
            firstTopDryer4Past = firstTopDryer4Past != null ? Math.Round((double)firstTopDryer4Past) : 0;
            cellFirstTopDryer4Past.AddParagraph($"{firstTopDryer4Past}");

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

            #endregion // 1st Top Dryer Position

            #region 2nd Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
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
            var secondTopDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 7 && c.StatusID == 3).Select(c => c.Age).Average();
            secondTopDryer1Past = secondTopDryer1Past != null ? Math.Round((double)secondTopDryer1Past) : 0;
            cell2ndTopDryer1Past.AddParagraph($"{secondTopDryer1Past}");

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
            var secondTopDryer2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 7 && c.StatusID == 3).Select(c => c.Age).Average();
            secondTopDryer2Past = secondTopDryer2Past != null ? Math.Round((double)secondTopDryer2Past) : 0;
            cell2ndTopDryer2Past.AddParagraph($"{secondTopDryer2Past}");

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
            var secondTopDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 7 && c.StatusID == 3).Select(c => c.Age).Average();
            secondTopDryer3Past = secondTopDryer3Past != null ? Math.Round((double)secondTopDryer3Past) : 0;
            cell2ndTopDryer3Past.AddParagraph($"{secondTopDryer3Past}");

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
            var secondTopDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 7 && c.StatusID == 3).Select(c => c.Age).Average();
            secondTopDryer4Past = secondTopDryer4Past != null ? Math.Round((double)secondTopDryer4Past) : 0;
            cell2ndTopDryer4Past.AddParagraph($"{secondTopDryer4Past}");

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
            spacer.Height = 12;
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
            var thirdTopDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 8 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdTopDryer1Past = thirdTopDryer1Past != null ? Math.Round((double)thirdTopDryer1Past) : 0;
            cell3rdTopDryer1Past.AddParagraph($"{thirdTopDryer1Past}");

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
            var thirdTopDryer2Past = db.Clothings.Where(c => c.PM_Number == 2  && c.PositionID == 8 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdTopDryer2Past = thirdTopDryer2Past != null ? Math.Round((double)thirdTopDryer2Past) : 0;
            cell3rdTopDryer2Past.AddParagraph($"{thirdTopDryer2Past}");

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
            var thirdTopDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 8 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdTopDryer3Past = thirdTopDryer3Past != null ? Math.Round((double)thirdTopDryer3Past) : 0;
            cell3rdTopDryer3Past.AddParagraph($"{thirdTopDryer3Past}");

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
            var thirdTopDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 8 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdTopDryer4Past = thirdTopDryer4Past != null ? Math.Round((double)thirdTopDryer4Past) : 0;
            cell3rdTopDryer4Past.AddParagraph($"{thirdTopDryer4Past}");

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
            spacer.Height = 12;
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
            var fourthTopDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 9 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthTopDryer1Past = fourthTopDryer1Past != null ? Math.Round((double)fourthTopDryer1Past) : 0;
            cell4thTopDryer1Past.AddParagraph($"{fourthTopDryer1Past}");

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
            var fourthTopDryer2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 9 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthTopDryer2Past = fourthTopDryer2Past != null ? Math.Round((double)fourthTopDryer2Past) : 0;
            cell4thTopDryer2Past.AddParagraph($"{fourthTopDryer2Past}");

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
            var fourthTopDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 9 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthTopDryer3Past = fourthTopDryer3Past != null ? Math.Round((double)fourthTopDryer3Past) : 0;
            cell4thTopDryer3Past.AddParagraph($"{fourthTopDryer3Past}");

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
            var fourthTopDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 9 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthTopDryer4Past = fourthTopDryer4Past != null ? Math.Round((double)fourthTopDryer4Past) : 0;
            cell4thTopDryer4Past.AddParagraph($"{fourthTopDryer4Past}"); ;

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

            #endregion // 4th Top Dryer Position

            #region 1st Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
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
            var firstBottomDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 10 && c.StatusID == 3).Select(c => c.Age).Average();
            firstBottomDryer1Past = firstBottomDryer1Past != null ? Math.Round((double)firstBottomDryer1Past) : 0;
            cell1stBottomDryer1Past.AddParagraph($"{firstBottomDryer1Past}");

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
            var firstBottomDryer2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 10 && c.StatusID == 3).Select(c => c.Age).Average();
            firstBottomDryer2Past = firstBottomDryer2Past != null ? Math.Round((double)firstBottomDryer2Past) : 0;
            cell1stBottomDryer2Past.AddParagraph($"{firstBottomDryer2Past}");

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
            var firstBottomDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 10 && c.StatusID == 3).Select(c => c.Age).Average();
            firstBottomDryer3Past = firstBottomDryer3Past != null ? Math.Round((double)firstBottomDryer3Past) : 0;
            cell1stBottomDryer3Past.AddParagraph($"{firstBottomDryer3Past}");

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
            var firstBottomDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 10 && c.StatusID == 3).Select(c => c.Age).Average();
            firstBottomDryer4Past = firstBottomDryer4Past != null ? Math.Round((double)firstBottomDryer4Past) : 0;
            cell1stBottomDryer4Past.AddParagraph($"{firstBottomDryer4Past}");

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

            #endregion // 1st Bottom Dryer Position

            #region 2nd Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
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
            var secondBottomDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 11 && c.StatusID == 3).Select(c => c.Age).Average();
            secondBottomDryer1Past = secondBottomDryer1Past != null ? Math.Round((double)secondBottomDryer1Past) : 0;
            cell2ndBottomDryer1Past.AddParagraph($"{secondBottomDryer1Past}");

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
            var secondBottomDryer2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 11 && c.StatusID == 3).Select(c => c.Age).Average();
            secondBottomDryer2Past = secondBottomDryer2Past != null ? Math.Round((double)secondBottomDryer2Past) : 0;
            cell2ndBottomDryer2Past.AddParagraph($"{secondBottomDryer2Past}");

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
            var secondBottomDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 11 && c.StatusID == 3).Select(c => c.Age).Average();
            secondBottomDryer3Past = secondBottomDryer3Past != null ? Math.Round((double)secondBottomDryer3Past) : 0;
            cell2ndBottomDryer3Past.AddParagraph($"{secondBottomDryer3Past}");

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
            var secondBottomDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 11 && c.StatusID == 3).Select(c => c.Age).Average();
            secondBottomDryer4Past = secondBottomDryer4Past != null ? Math.Round((double)secondBottomDryer4Past) : 0;
            cell2ndBottomDryer4Past.AddParagraph($"{secondBottomDryer4Past}");

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

            #endregion // 2nd Bottom Dryer Position

            #region 3rd Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
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
            var thirdBottomDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 12 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdBottomDryer1Past = thirdBottomDryer1Past != null ? Math.Round((double)thirdBottomDryer1Past) : 0;
            cell3rdBottomDryer1Past.AddParagraph($"{thirdBottomDryer1Past}");

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
            var thirdBottomDryer2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 12 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdBottomDryer2Past = thirdBottomDryer2Past != null ? Math.Round((double)thirdBottomDryer2Past) : 0;
            cell3rdBottomDryer2Past.AddParagraph($"{thirdBottomDryer2Past}");

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
            var thirdBottomDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 12 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdBottomDryer3Past = thirdBottomDryer3Past != null ? Math.Round((double)thirdBottomDryer3Past) : 0;
            cell3rdBottomDryer3Past.AddParagraph($"{thirdBottomDryer3Past}");

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
            var thirdBottomDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 12 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdBottomDryer4Past = thirdBottomDryer4Past != null ? Math.Round((double)thirdBottomDryer4Past) : 0;
            cell3rdBottomDryer4Past.AddParagraph($"{thirdBottomDryer4Past}");

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

            #endregion // 3rd Bottom Dryer Position

            #region 4th Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
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
            var fourthBottomDryer1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 13 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthBottomDryer1Past = fourthBottomDryer1Past != null ? Math.Round((double)fourthBottomDryer1Past) : 0;
            cell4thBottomDryer1Past.AddParagraph($"{fourthBottomDryer1Past}");

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
            var fourthBottomDryer2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 13 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthBottomDryer2Past = fourthBottomDryer2Past != null ? Math.Round((double)fourthBottomDryer2Past) : 0;
            cell4thBottomDryer2Past.AddParagraph($"{fourthBottomDryer2Past}");

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
            var fourthBottomDryer3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 13 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthBottomDryer3Past = fourthBottomDryer3Past != null ? Math.Round((double)fourthBottomDryer3Past) : 0;
            cell4thBottomDryer3Past.AddParagraph($"{fourthBottomDryer3Past}");

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
            var fourthBottomDryer4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 13 && c.StatusID == 3).Select(c => c.Age).Average();
            fourthBottomDryer4Past = fourthBottomDryer4Past != null ? Math.Round((double)fourthBottomDryer4Past) : 0;
            cell4thBottomDryer4Past.AddParagraph($"{fourthBottomDryer4Past}");

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

            #endregion // 4th Bottom Dryer Position

            #region Breast Roll Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
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
            var breast1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 14 && c.StatusID == 3).Select(c => c.Age).Average();
            breast1Past = breast1Past != null ? Math.Round((double)breast1Past) : 0;
            cellBreast1Past.AddParagraph($"{breast1Past}");

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
            var breast2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 14 && c.StatusID == 3).Select(c => c.Age).Average();
            breast2Past = breast2Past != null ? Math.Round((double)breast2Past) : 0;
            cellBreast2Past.AddParagraph($"{breast2Past}");

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
            var breast3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 14 && c.StatusID == 3).Select(c => c.Age).Average();
            breast3Past = breast3Past != null ? Math.Round((double)breast3Past) : 0;
            cellBreast3Past.AddParagraph($"{breast3Past}");

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
            var breast4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 14 && c.StatusID == 3).Select(c => c.Age).Average();
            breast4Past = breast4Past != null ? Math.Round((double)breast4Past) : 0;
            cellBreast4Past.AddParagraph($"{breast4Past}");

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

            #endregion // Breast Roll Position

            #region Dandy Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingDandy1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 1);
            var clothingDandy2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 2);
            var clothingDandy3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 3);
            var clothingDandy4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 4);
            var rowDandy = weeklyPMTable.AddRow();
            rowDandy.Height = 10;
            var rowDandy2 = weeklyPMTable.AddRow();
            rowDandy2.Height = 10;

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
            var dandy1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 15 && c.StatusID == 3).Select(c => c.Age).Average();
            dandy1Past = dandy1Past != null ? Math.Round((double)dandy1Past) : 0;
            cellDandy1Past.AddParagraph($"{dandy1Past}");

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
            var dandy2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 15 && c.StatusID == 3).Select(c => c.Age).Average();
            dandy2Past = dandy2Past != null ? Math.Round((double)dandy2Past) : 0;
            cellDandy2Past.AddParagraph($"{dandy2Past}");

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
            var dandy3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 15 && c.StatusID == 3).Select(c => c.Age).Average();
            dandy3Past = dandy3Past != null ? Math.Round((double)dandy3Past) : 0;
            cellDandy3Past.AddParagraph($"{dandy3Past}");

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
            var dandy4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 15 && c.StatusID == 3).Select(c => c.Age).Average();
            dandy4Past = dandy4Past != null ? Math.Round((double)dandy4Past) : 0;
            cellDandy4Past.AddParagraph($"{dandy4Past}");

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

            #endregion // Dandy Roll Position

            #region Lumpbreaker Position

            var clothingLumpbreaker1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 1);
            var clothingLumpbreaker2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 2);
            var clothingLumpbreaker3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 3);
            var clothingLumpbreaker4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 4);
            var rowLumpbreaker = weeklyPMTable.AddRow();
            rowLumpbreaker.Height = 10;
            var rowLumpbreaker2 = weeklyPMTable.AddRow();
            rowLumpbreaker2.Height = 10;

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
            var lumpbreaker1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 16 && c.StatusID == 3).Select(c => c.Age).Average();
            lumpbreaker1Past = lumpbreaker1Past != null ? Math.Round((double)lumpbreaker1Past) : 0;
            cellLumpbreaker1Past.AddParagraph($"{lumpbreaker1Past}");

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
            var lumpbreaker2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 16 && c.StatusID == 3).Select(c => c.Age).Average();
            lumpbreaker2Past = lumpbreaker2Past != null ? Math.Round((double)lumpbreaker2Past) : 0;
            cellLumpbreaker2Past.AddParagraph($"{lumpbreaker2Past}");

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
            var lumpbreaker3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 16 && c.StatusID == 3).Select(c => c.Age).Average();
            lumpbreaker3Past = lumpbreaker3Past != null ? Math.Round((double)lumpbreaker3Past) : 0;
            cellLumpbreaker3Past.AddParagraph($"{lumpbreaker3Past}");

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
            var lumpbreaker4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 16 && c.StatusID == 3).Select(c => c.Age).Average();
            lumpbreaker4Past = lumpbreaker4Past != null ? Math.Round((double)lumpbreaker4Past) : 0;
            cellLumpbreaker4Past.AddParagraph($"{lumpbreaker4Past}");

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

            #endregion // Lumpbreaker Roll Position

            #region Suction Pickup Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSuction1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 1);
            var clothingSuction2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 2);
            var clothingSuction3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 3);
            var clothingSuction4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 4);
            var rowSuction = weeklyPMTable.AddRow();
            rowSuction.Height = 10;
            var rowSuction2 = weeklyPMTable.AddRow();
            rowSuction2.Height = 10;

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
            var suction1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 17 && c.StatusID == 3).Select(c => c.Age).Average();
            suction1Past = suction1Past != null ? Math.Round((double)suction1Past) : 0;
            cellSuction1Past.AddParagraph($"{suction1Past}");

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
            var suction2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 17 && c.StatusID == 3).Select(c => c.Age).Average();
            suction2Past = suction2Past != null ? Math.Round((double)suction2Past) : 0;
            cellSuction2Past.AddParagraph($"{suction2Past}");

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
            var suction3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 17 && c.StatusID == 3).Select(c => c.Age).Average();
            suction3Past = suction3Past != null ? Math.Round((double)suction3Past) : 0;
            cellSuction3Past.AddParagraph($"{suction3Past}");

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
            var suction4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 17 && c.StatusID == 3).Select(c => c.Age).Average();
            suction4Past = suction4Past != null ? Math.Round((double)suction4Past) : 0;
            cellSuction4Past.AddParagraph($"{suction4Past}");

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
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stPressTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 1);
            var clothing1stPressTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 2);
            var clothing1stPressTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 3);
            var clothing1stPressTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 4);
            var row1stPressTop = weeklyPMTable.AddRow();
            row1stPressTop.Height = 10;
            var row1stPressTop2 = weeklyPMTable.AddRow();
            row1stPressTop2.Height = 10;

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
            var firstPressTop1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 18 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressTop1Past = firstPressTop1Past != null ? Math.Round((double)firstPressTop1Past) : 0;
            cell1stPressTop1Past.AddParagraph($"{firstPressTop1Past}");

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
            var firstPressTop2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 18 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressTop2Past = firstPressTop2Past != null ? Math.Round((double)firstPressTop2Past) : 0;
            cell1stPressTop2Past.AddParagraph($"{firstPressTop2Past}");

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
            var firstPressTop3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 18 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressTop3Past = firstPressTop3Past != null ? Math.Round((double)firstPressTop3Past) : 0;
            cell1stPressTop3Past.AddParagraph($"{firstPressTop3Past}");

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
            var firstPressTop4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 18 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressTop4Past = firstPressTop4Past != null ? Math.Round((double)firstPressTop4Past) : 0;
            cell1stPressTop4Past.AddParagraph($"{firstPressTop4Past}");

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
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stPressBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 1);
            var clothing1stPressBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 2);
            var clothing1stPressBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 3);
            var clothing1stPressBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 4);
            var row1stPressBottom = weeklyPMTable.AddRow();
            row1stPressBottom.Height = 10;
            var row1stPressBottom2 = weeklyPMTable.AddRow();
            row1stPressBottom2.Height = 10;

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
            var firstPressBottom1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 19 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressBottom1Past = firstPressBottom1Past != null ? Math.Round((double)firstPressBottom1Past) : 0;
            cell1stPressBottom1Past.AddParagraph($"{firstPressBottom1Past}");

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

            #region 1st Press Bottom PM 2
            // ----- 1st Press Bottom Position PM 2 ----- //
            var cell1stPressBottom2Past = row1stPressBottom.Cells[4];
            cell1stPressBottom2Past.Format.Font.Size = 6.5;
            cell1stPressBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 19 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressBottom2Past = firstPressBottom2Past != null ? Math.Round((double)firstPressBottom2Past) : 0;
            cell1stPressBottom2Past.AddParagraph($"{firstPressBottom2Past}");

            var cell1stPressBottom2Current = row1stPressBottom.Cells[5];
            cell1stPressBottom2Current.Format.Font.Size = 6.5;
            cell1stPressBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom2Age = 0;
            if (clothing1stPressBottom2 != null)
            {
                firstPressBottom2Age = clothing1stPressBottom2.Age != null ? Convert.ToInt32(clothing1stPressBottom2.Age) : 0;
            }
            cell1stPressBottom2Current.AddParagraph($"{firstPressBottom2Age}");

            var cell1stPressBottom2Goal = row1stPressBottom.Cells[6];
            cell1stPressBottom2Goal.Format.Font.Size = 6.5;
            cell1stPressBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 19);
            cell1stPressBottom2Goal.AddParagraph($"{firstPressBottom2Goal.Goal1}");

            var cell1stPressBottom2Info = row1stPressBottom2.Cells[4];
            cell1stPressBottom2Info.Format.Font.Size = 6.5;
            cell1stPressBottom2Info.MergeRight = 2;
            cell1stPressBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom2Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressBottom2 != null)
            {
                cell1stPressBottom2Info.AddParagraph($"{clothing1stPressBottom2.Manufacturer} {clothing1stPressBottom2.Serial_Number}");
            }
            else cell1stPressBottom2Info.AddParagraph("NA");
            #endregion

            #region 1st Press Bottom PM 3
            // ----- 1st Press Bottom Position PM 3 ----- //
            var cell1stPressBottom3Past = row1stPressBottom.Cells[7];
            cell1stPressBottom3Past.Format.Font.Size = 6.5;
            cell1stPressBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 19 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressBottom3Past = firstPressBottom3Past != null ? Math.Round((double)firstPressBottom3Past) : 0;
            cell1stPressBottom3Past.AddParagraph($"{firstPressBottom3Past}");

            var cell1stPressBottom3Current = row1stPressBottom.Cells[8];
            cell1stPressBottom3Current.Format.Font.Size = 6.5;
            cell1stPressBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom3Age = 0;
            if (clothing1stPressBottom3 != null)
            {
                firstPressBottom3Age = clothing1stPressBottom3.Age != null ? Convert.ToInt32(clothing1stPressBottom3.Age) : 0;
            }
            cell1stPressBottom3Current.AddParagraph($"{firstPressBottom3Age}");

            var cell1stPressBottom3Goal = row1stPressBottom.Cells[9];
            cell1stPressBottom3Goal.Format.Font.Size = 6.5;
            cell1stPressBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 19);
            cell1stPressBottom3Goal.AddParagraph($"{firstPressBottom3Goal.Goal1}");

            var cell1stPressBottom3Info = row1stPressBottom2.Cells[7];
            cell1stPressBottom3Info.Format.Font.Size = 6.5;
            cell1stPressBottom3Info.MergeRight = 2;
            cell1stPressBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom3Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressBottom3 != null)
            {
                cell1stPressBottom3Info.AddParagraph($"{clothing1stPressBottom3.Manufacturer} {clothing1stPressBottom3.Serial_Number}");
            }
            else cell1stPressBottom3Info.AddParagraph("NA");
            #endregion

            #region 1st Press Bottom PM 4
            // ----- 1st Press Bottom Position PM 4 ----- //
            var cell1stPressBottom4Past = row1stPressBottom.Cells[10];
            cell1stPressBottom4Past.Format.Font.Size = 6.5;
            cell1stPressBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 19 && c.StatusID == 3).Select(c => c.Age).Average();
            firstPressBottom4Past = firstPressBottom4Past != null ? Math.Round((double)firstPressBottom4Past) : 0;
            cell1stPressBottom4Past.AddParagraph($"{firstPressBottom4Past}");

            var cell1stPressBottom4Current = row1stPressBottom.Cells[11];
            cell1stPressBottom4Current.Format.Font.Size = 6.5;
            cell1stPressBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom4Age = 0;
            if (clothing1stPressBottom4 != null)
            {
                firstPressBottom4Age = clothing1stPressBottom4.Age != null ? Convert.ToInt32(clothing1stPressBottom4.Age) : 0;
            }
            cell1stPressBottom4Current.AddParagraph($"{firstPressBottom4Age}");

            var cell1stPressBottom4Goal = row1stPressBottom.Cells[12];
            cell1stPressBottom4Goal.Format.Font.Size = 6.5;
            cell1stPressBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            var firstPressBottom4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 19);
            cell1stPressBottom4Goal.AddParagraph($"{firstPressBottom4Goal.Goal1}");

            var cell1stPressBottom4Info = row1stPressBottom2.Cells[10];
            cell1stPressBottom4Info.Format.Font.Size = 6.5;
            cell1stPressBottom4Info.MergeRight = 2;
            cell1stPressBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom4Info.Shading.Color = Colors.LightBlue;
            if (clothing1stPressBottom4 != null)
            {
                cell1stPressBottom4Info.AddParagraph($"{clothing1stPressBottom4.Manufacturer} {clothing1stPressBottom4.Serial_Number}");
            }
            else cell1stPressBottom4Info.AddParagraph("NA");
            #endregion

            #endregion // 1st Press Bottom Position

            #region 2nd Press Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndPressTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 1);
            var clothing2ndPressTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 2);
            var clothing2ndPressTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 3);
            var clothing2ndPressTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 4);
            var row2ndPressTop = weeklyPMTable.AddRow();
            row2ndPressTop.Height = 10;
            var row2ndPressTop2 = weeklyPMTable.AddRow();
            row2ndPressTop2.Height = 10;

            var cell2ndPressTop = row2ndPressTop.Cells[0];
            cell2ndPressTop.Format.Font.Size = 6.5;
            cell2ndPressTop.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop.MergeDown = 1;
            cell2ndPressTop.AddParagraph("2ND PRESS TOP");

            #region 2nd Press Top PM 1
            // ----- 2nd Press Top Position PM 1 ----- //
            var cell2ndPressTop1Past = row2ndPressTop.Cells[1];
            cell2ndPressTop1Past.Format.Font.Size = 6.5;
            cell2ndPressTop1Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 20 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressTop1Past = secondPressTop1Past != null ? Math.Round((double)secondPressTop1Past) : 0;
            cell2ndPressTop1Past.AddParagraph($"{secondPressTop1Past}");

            var cell2ndPressTop1Current = row2ndPressTop.Cells[2];
            cell2ndPressTop1Current.Format.Font.Size = 6.5;
            cell2ndPressTop1Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop1Age = 0;
            if (clothing2ndPressTop1 != null)
            {
                secondPressTop1Age = clothing2ndPressTop1.Age != null ? Convert.ToInt32(clothing2ndPressTop1.Age) : 0;
            }
            cell2ndPressTop1Current.AddParagraph($"{secondPressTop1Age}");

            var cell2ndPressTop1Goal = row2ndPressTop.Cells[3];
            cell2ndPressTop1Goal.Format.Font.Size = 6.5;
            cell2ndPressTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressToplGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 20);
            cell2ndPressTop1Goal.AddParagraph($"{secondPressToplGoal.Goal1}");

            var cell2ndPressTop1Info = row2ndPressTop2.Cells[1];
            cell2ndPressTop1Info.Format.Font.Size = 6.5;
            cell2ndPressTop1Info.MergeRight = 2;
            cell2ndPressTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop1Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressTop1 != null)
            {
                cell2ndPressTop1Info.AddParagraph($"{clothing2ndPressTop1.Manufacturer} {clothing2ndPressTop1.Serial_Number}");
            }
            else cell2ndPressTop1Info.AddParagraph("NA");
            #endregion

            #region 2nd Press Top PM 2
            // ----- 2nd Press Top Position PM 2 ----- //
            var cell2ndPressTop2Past = row2ndPressTop.Cells[4];
            cell2ndPressTop2Past.Format.Font.Size = 6.5;
            cell2ndPressTop2Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 20 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressTop2Past = secondPressTop2Past != null ? Math.Round((double)secondPressTop2Past) : 0;
            cell2ndPressTop2Past.AddParagraph($"{secondPressTop2Past}");

            var cell2ndPressTop2Current = row2ndPressTop.Cells[5];
            cell2ndPressTop2Current.Format.Font.Size = 6.5;
            cell2ndPressTop2Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop2Age = 0;
            if (clothing2ndPressTop2 != null)
            {
                secondPressTop2Age = clothing2ndPressTop2.Age != null ? Convert.ToInt32(clothing2ndPressTop2.Age) : 0;
            }
            cell2ndPressTop2Current.AddParagraph($"{secondPressTop2Age}");

            var cell2ndPressTop2Goal = row2ndPressTop.Cells[6];
            cell2ndPressTop2Goal.Format.Font.Size = 6.5;
            cell2ndPressTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 20);
            cell2ndPressTop2Goal.AddParagraph($"{secondPressTop2Goal.Goal1}");

            var cell2ndPressTop2Info = row2ndPressTop2.Cells[4];
            cell2ndPressTop2Info.Format.Font.Size = 6.5;
            cell2ndPressTop2Info.MergeRight = 2;
            cell2ndPressTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop2Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressTop2 != null)
            {
                cell2ndPressTop2Info.AddParagraph($"{clothing2ndPressTop2.Manufacturer} {clothing2ndPressTop2.Serial_Number}");
            }
            else cell2ndPressTop2Info.AddParagraph("NA");
            #endregion

            #region 2nd Press Top PM 3
            // -----2nd Press Top Position PM 3---- - //
            var cell2ndPressTop3Past = row2ndPressTop.Cells[7];
            cell2ndPressTop3Past.Format.Font.Size = 6.5;
            cell2ndPressTop3Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 20 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressTop3Past = secondPressTop3Past != null ? Math.Round((double)secondPressTop3Past) : 0;
            cell2ndPressTop3Past.AddParagraph($"{secondPressTop3Past}");

            var cell2ndPressTop3Current = row2ndPressTop.Cells[8];
            cell2ndPressTop3Current.Format.Font.Size = 6.5;
            cell2ndPressTop3Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop3Age = 0;
            if (clothing2ndPressTop3 != null)
            {
                secondPressTop3Age = clothing2ndPressTop3.Age != null ? Convert.ToInt32(clothing2ndPressTop3.Age) : 0;
            }
            cell2ndPressTop3Current.AddParagraph($"{secondPressTop3Age}");

            var cell2ndPressTop3Goal = row2ndPressTop.Cells[9];
            cell2ndPressTop3Goal.Format.Font.Size = 6.5;
            cell2ndPressTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 20);
            cell2ndPressTop3Goal.AddParagraph($"{secondPressTop3Goal.Goal1}");

            var cell2ndPressTop3Info = row2ndPressTop2.Cells[7];
            cell2ndPressTop3Info.Format.Font.Size = 6.5;
            cell2ndPressTop3Info.MergeRight = 2;
            cell2ndPressTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop3Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressTop3 != null)
            {
                cell2ndPressTop3Info.AddParagraph($"{clothing2ndPressTop3.Manufacturer} {clothing2ndPressTop3.Serial_Number}");
            }
            else cell2ndPressTop3Info.AddParagraph("NA");
            #endregion

            #region 2nd Press Top PM 4
            // -----2nd Press Top Position PM 4---- - //
            var cell2ndPressTop4Past = row2ndPressTop.Cells[10];
            cell2ndPressTop4Past.Format.Font.Size = 6.5;
            cell2ndPressTop4Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 20 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressTop4Past = secondPressTop4Past != null ? Math.Round((double)secondPressTop4Past) : 0;
            cell2ndPressTop4Past.AddParagraph($"{secondPressTop4Past}");

            var cell2ndPressTop4Current = row2ndPressTop.Cells[11];
            cell2ndPressTop4Current.Format.Font.Size = 6.5;
            cell2ndPressTop4Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop4Age = 0;
            if (clothing2ndPressTop4 != null)
            {
                secondPressTop4Age = clothing2ndPressTop4.Age != null ? Convert.ToInt32(clothing2ndPressTop4.Age) : 0;
            }
            cell2ndPressTop4Current.AddParagraph($"{secondPressTop4Age}");

            var cell2ndPressTop4Goal = row2ndPressTop.Cells[12];
            cell2ndPressTop4Goal.Format.Font.Size = 6.5;
            cell2ndPressTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressTop4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 20);
            cell2ndPressTop4Goal.AddParagraph($"{secondPressTop4Goal.Goal1}");

            var cell2ndPressTop4Info = row2ndPressTop2.Cells[10];
            cell2ndPressTop4Info.Format.Font.Size = 6.5;
            cell2ndPressTop4Info.MergeRight = 2;
            cell2ndPressTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop4Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressTop4 != null)
            {
                cell2ndPressTop4Info.AddParagraph($"{clothing2ndPressTop4.Manufacturer} {clothing2ndPressTop4.Serial_Number}");
            }
            else cell2ndPressTop4Info.AddParagraph("NA");
            #endregion

            #endregion // 2nd Press Top Position

            #region 2nd Press Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndPressBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 1);
            var clothing2ndPressBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 2);
            var clothing2ndPressBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 3);
            var clothing2ndPressBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 4);
            var row2ndPressBottom = weeklyPMTable.AddRow();
            row2ndPressBottom.Height = 10;
            var row2ndPressBottom2 = weeklyPMTable.AddRow();
            row2ndPressBottom2.Height = 10;

            var cell2ndPressBottom = row2ndPressBottom.Cells[0];
            cell2ndPressBottom.Format.Font.Size = 6.5;
            cell2ndPressBottom.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom.MergeDown = 1;
            cell2ndPressBottom.AddParagraph("2ND PRESS BOTTOM");

            #region 2nd Press Bottom PM 1
            // ----- 2nd Press Bottom Position PM 1 ----- //
            var cell2ndPressBottom1Past = row2ndPressBottom.Cells[1];
            cell2ndPressBottom1Past.Format.Font.Size = 6.5;
            cell2ndPressBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 21 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressBottom1Past = secondPressBottom1Past != null ? Math.Round((double)secondPressBottom1Past) : 0;
            cell2ndPressBottom1Past.AddParagraph($"{secondPressBottom1Past}");

            var cell2ndPressBottom1Current = row2ndPressBottom.Cells[2];
            cell2ndPressBottom1Current.Format.Font.Size = 6.5;
            cell2ndPressBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom1Age = 0;
            if (clothing2ndPressBottom1 != null)
            {
                secondPressBottom1Age = clothing2ndPressBottom1.Age != null ? Convert.ToInt32(clothing2ndPressBottom1.Age) : 0;
            }
            cell2ndPressBottom1Current.AddParagraph($"{secondPressBottom1Age}");

            var cell2ndPressBottom1Goal = row2ndPressBottom.Cells[3];
            cell2ndPressBottom1Goal.Format.Font.Size = 6.5;
            cell2ndPressBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottomlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 21);
            cell2ndPressBottom1Goal.AddParagraph($"{secondPressBottomlGoal.Goal1}");

            var cell2ndPressBottom1Info = row2ndPressBottom2.Cells[1];
            cell2ndPressBottom1Info.Format.Font.Size = 6.5;
            cell2ndPressBottom1Info.MergeRight = 2;
            cell2ndPressBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom1Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressBottom1 != null)
            {
                cell2ndPressBottom1Info.AddParagraph($"{clothing2ndPressBottom1.Manufacturer} {clothing2ndPressBottom1.Serial_Number}");
            }
            else cell2ndPressBottom1Info.AddParagraph("NA");
            #endregion

            #region 2nd Press Bottom PM 2
            // ----- 2nd Press Bottom Position PM 2 ----- //
            var cell2ndPressBottom2Past = row2ndPressBottom.Cells[4];
            cell2ndPressBottom2Past.Format.Font.Size = 6.5;
            cell2ndPressBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 21 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressBottom2Past = secondPressBottom2Past != null ? Math.Round((double)secondPressBottom2Past) : 0;
            cell2ndPressBottom2Past.AddParagraph($"{secondPressBottom2Past}");

            var cell2ndPressBottom2Current = row2ndPressBottom.Cells[5];
            cell2ndPressBottom2Current.Format.Font.Size = 6.5;
            cell2ndPressBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom2Age = 0;
            if (clothing2ndPressBottom2 != null)
            {
                secondPressBottom2Age = clothing2ndPressBottom2.Age != null ? Convert.ToInt32(clothing2ndPressBottom2.Age) : 0;
            }
            cell2ndPressBottom2Current.AddParagraph($"{secondPressBottom2Age}");

            var cell2ndPressBottom2Goal = row2ndPressBottom.Cells[6];
            cell2ndPressBottom2Goal.Format.Font.Size = 6.5;
            cell2ndPressBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 21);
            cell2ndPressBottom2Goal.AddParagraph($"{secondPressBottom2Goal.Goal1}");

            var cell2ndPressBottom2Info = row2ndPressBottom2.Cells[4];
            cell2ndPressBottom2Info.Format.Font.Size = 6.5;
            cell2ndPressBottom2Info.MergeRight = 2;
            cell2ndPressBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom2Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressBottom2 != null)
            {
                cell2ndPressBottom2Info.AddParagraph($"{clothing2ndPressBottom2.Manufacturer} {clothing2ndPressBottom2.Serial_Number}");
            }
            else cell2ndPressBottom2Info.AddParagraph("NA");
            #endregion

            #region 2nd Press Bottom PM 3
            // ----- 2nd Press Bottom Position PM 3 ----- //
            var cell2ndPressBottom3Past = row2ndPressBottom.Cells[7];
            cell2ndPressBottom3Past.Format.Font.Size = 6.5;
            cell2ndPressBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 21 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressBottom3Past = secondPressBottom3Past != null ? Math.Round((double)secondPressBottom3Past) : 0;
            cell2ndPressBottom3Past.AddParagraph($"{secondPressBottom3Past}");

            var cell2ndPressBottom3Current = row2ndPressBottom.Cells[8];
            cell2ndPressBottom3Current.Format.Font.Size = 6.5;
            cell2ndPressBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom3Age = 0;
            if (clothing2ndPressBottom3 != null)
            {
                secondPressBottom3Age = clothing2ndPressBottom3.Age != null ? Convert.ToInt32(clothing2ndPressBottom3.Age) : 0;
            }
            cell2ndPressBottom3Current.AddParagraph($"{secondPressBottom3Age}");

            var cell2ndPressBottom3Goal = row2ndPressBottom.Cells[9];
            cell2ndPressBottom3Goal.Format.Font.Size = 6.5;
            cell2ndPressBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 21);
            cell2ndPressBottom3Goal.AddParagraph($"{secondPressBottom3Goal.Goal1}");

            var cell2ndPressBottom3Info = row2ndPressBottom2.Cells[7];
            cell2ndPressBottom3Info.Format.Font.Size = 6.5;
            cell2ndPressBottom3Info.MergeRight = 2;
            cell2ndPressBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom3Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressBottom3 != null)
            {
                cell2ndPressBottom3Info.AddParagraph($"{clothing2ndPressBottom3.Manufacturer} {clothing2ndPressBottom3.Serial_Number}");
            }
            else cell2ndPressBottom3Info.AddParagraph("NA");
            #endregion

            #region 2nd Press Bottom PM 4
            // ----- 2nd Press Bottom Position PM 4 ----- //
            var cell2ndPressBottom4Past = row2ndPressBottom.Cells[10];
            cell2ndPressBottom4Past.Format.Font.Size = 6.5;
            cell2ndPressBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 21 && c.StatusID == 3).Select(c => c.Age).Average();
            secondPressBottom4Past = secondPressBottom4Past != null ? Math.Round((double)secondPressBottom4Past) : 0;
            cell2ndPressBottom4Past.AddParagraph($"{secondPressBottom4Past}");

            var cell2ndPressBottom4Current = row2ndPressBottom.Cells[11];
            cell2ndPressBottom4Current.Format.Font.Size = 6.5;
            cell2ndPressBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom4Age = 0;
            if (clothing2ndPressBottom4 != null)
            {
                secondPressBottom4Age = clothing2ndPressBottom4.Age != null ? Convert.ToInt32(clothing2ndPressBottom4.Age) : 0;
            }
            cell2ndPressBottom4Current.AddParagraph($"{secondPressBottom4Age}");

            var cell2ndPressBottom4Goal = row2ndPressBottom.Cells[12];
            cell2ndPressBottom4Goal.Format.Font.Size = 6.5;
            cell2ndPressBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondPressBottom4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 21);
            cell2ndPressBottom4Goal.AddParagraph($"{secondPressBottom4Goal.Goal1}");

            var cell2ndPressBottom4Info = row2ndPressBottom2.Cells[10];
            cell2ndPressBottom4Info.Format.Font.Size = 6.5;
            cell2ndPressBottom4Info.MergeRight = 2;
            cell2ndPressBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom4Info.Shading.Color = Colors.LightBlue;
            if (clothing2ndPressBottom4 != null)
            {
                cell2ndPressBottom4Info.AddParagraph($"{clothing2ndPressBottom4.Manufacturer} {clothing2ndPressBottom4.Serial_Number}");
            }
            else cell2ndPressBottom4Info.AddParagraph("NA");
            #endregion

            #endregion // 2nd Press Bottom Position

            #region 3rd Press Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdPressTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 1);
            var clothing3rdPressTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 2);
            var clothing3rdPressTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 3);
            var clothing3rdPressTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 4);
            var row3rdPressTop = weeklyPMTable.AddRow();
            row3rdPressTop.Height = 10;
            var row3rdPressTop2 = weeklyPMTable.AddRow();
            row3rdPressTop2.Height = 10;

            var cell3rdPressTop = row3rdPressTop.Cells[0];
            cell3rdPressTop.Format.Font.Size = 6.5;
            cell3rdPressTop.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop.MergeDown = 1;
            cell3rdPressTop.AddParagraph("3RD PRESS TOP");

            #region 3rd Press Top PM 1
            // ----- 3rd Press Top Position PM 1 ----- //
            var cell3rdPressTop1Past = row3rdPressTop.Cells[1];
            cell3rdPressTop1Past.Format.Font.Size = 6.5;
            cell3rdPressTop1Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 22 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressTop1Past = thirdPressTop1Past != null ? Math.Round((double)thirdPressTop1Past) : 0;
            cell3rdPressTop1Past.AddParagraph($"{thirdPressTop1Past}");

            var cell3rdPressTop1Current = row3rdPressTop.Cells[2];
            cell3rdPressTop1Current.Format.Font.Size = 6.5;
            cell3rdPressTop1Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop1Age = 0;
            if (clothing3rdPressTop1 != null)
            {
                thirdPressTop1Age = clothing3rdPressTop1.Age != null ? Convert.ToInt32(clothing3rdPressTop1.Age) : 0;
            }
            cell3rdPressTop1Current.AddParagraph($"{thirdPressTop1Age}");

            var cell3rdPressTop1Goal = row3rdPressTop.Cells[3];
            cell3rdPressTop1Goal.Format.Font.Size = 6.5;
            cell3rdPressTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressToplGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 22);
            cell3rdPressTop1Goal.AddParagraph($"{thirdPressToplGoal.Goal1}");

            var cell3rdPressTop1Info = row3rdPressTop2.Cells[1];
            cell3rdPressTop1Info.Format.Font.Size = 6.5;
            cell3rdPressTop1Info.MergeRight = 2;
            cell3rdPressTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop1Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressTop1 != null)
            {
                cell3rdPressTop1Info.AddParagraph($"{clothing3rdPressTop1.Manufacturer} {clothing3rdPressTop1.Serial_Number}");
            }
            else cell3rdPressTop1Info.AddParagraph("NA");
            #endregion

            #region 3rd Press Top PM 2
            // ----- 3rd Press Top Position PM 2 ----- //
            var cell3rdPressTop2Past = row3rdPressTop.Cells[4];
            cell3rdPressTop2Past.Format.Font.Size = 6.5;
            cell3rdPressTop2Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 22 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressTop2Past = thirdPressTop2Past != null ? Math.Round((double)thirdPressTop2Past) : 0;
            cell3rdPressTop2Past.AddParagraph($"{thirdPressTop2Past}");

            var cell3rdPressTop2Current = row3rdPressTop.Cells[5];
            cell3rdPressTop2Current.Format.Font.Size = 6.5;
            cell3rdPressTop2Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop2Age = 0;
            if (clothing3rdPressTop2 != null)
            {
                thirdPressTop2Age = clothing3rdPressTop2.Age != null ? Convert.ToInt32(clothing3rdPressTop2.Age) : 0;
            }
            cell3rdPressTop2Current.AddParagraph($"{thirdPressTop2Age}");

            var cell3rdPressTop2Goal = row3rdPressTop.Cells[6];
            cell3rdPressTop2Goal.Format.Font.Size = 6.5;
            cell3rdPressTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 22);
            cell3rdPressTop2Goal.AddParagraph($"{thirdPressTop2Goal.Goal1}");

            var cell3rdPressTop2Info = row3rdPressTop2.Cells[4];
            cell3rdPressTop2Info.Format.Font.Size = 6.5;
            cell3rdPressTop2Info.MergeRight = 2;
            cell3rdPressTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop2Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressTop2 != null)
            {
                cell3rdPressTop2Info.AddParagraph($"{clothing3rdPressTop2.Manufacturer} {clothing3rdPressTop2.Serial_Number}");
            }
            else cell3rdPressTop2Info.AddParagraph("NA");
            #endregion

            #region 3rd Press Top PM 3
            // ----- 3rd Press Top Position PM 3 ----- //
            var cell3rdPressTop3Past = row3rdPressTop.Cells[7];
            cell3rdPressTop3Past.Format.Font.Size = 6.5;
            cell3rdPressTop3Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 22 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressTop3Past = thirdPressTop3Past != null ? Math.Round((double)thirdPressTop3Past) : 0;
            cell3rdPressTop3Past.AddParagraph($"{thirdPressTop3Past}");

            var cell3rdPressTop3Current = row3rdPressTop.Cells[8];
            cell3rdPressTop3Current.Format.Font.Size = 6.5;
            cell3rdPressTop3Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop3Age = 0;
            if (clothing3rdPressTop3 != null)
            {
                thirdPressTop3Age = clothing3rdPressTop3.Age != null ? Convert.ToInt32(clothing3rdPressTop3.Age) : 0;
            }
            cell3rdPressTop3Current.AddParagraph($"{thirdPressTop3Age}");

            var cell3rdPressTop3Goal = row3rdPressTop.Cells[9];
            cell3rdPressTop3Goal.Format.Font.Size = 6.5;
            cell3rdPressTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 22);
            cell3rdPressTop3Goal.AddParagraph($"{thirdPressTop3Goal.Goal1}");

            var cell3rdPressTop3Info = row3rdPressTop2.Cells[7];
            cell3rdPressTop3Info.Format.Font.Size = 6.5;
            cell3rdPressTop3Info.MergeRight = 2;
            cell3rdPressTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop3Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressTop3 != null)
            {
                cell3rdPressTop3Info.AddParagraph($"{clothing3rdPressTop3.Manufacturer} {clothing3rdPressTop3.Serial_Number}");
            }
            else cell3rdPressTop3Info.AddParagraph("NA");
            #endregion

            #region 3rd Press Top PM 4
            // ----- 3rd Press Top Position PM 4 ----- //
            var cell3rdPressTop4Past = row3rdPressTop.Cells[10];
            cell3rdPressTop4Past.Format.Font.Size = 6.5;
            cell3rdPressTop4Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 22 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressTop4Past = thirdPressTop4Past != null ? Math.Round((double)thirdPressTop4Past) : 0;
            cell3rdPressTop4Past.AddParagraph($"{thirdPressTop4Past}");

            var cell3rdPressTop4Current = row3rdPressTop.Cells[11];
            cell3rdPressTop4Current.Format.Font.Size = 6.5;
            cell3rdPressTop4Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop4Age = 0;
            if (clothing3rdPressTop4 != null)
            {
                thirdPressTop4Age = clothing3rdPressTop4.Age != null ? Convert.ToInt32(clothing3rdPressTop4.Age) : 0;
            }
            cell3rdPressTop4Current.AddParagraph($"{thirdPressTop4Age}");

            var cell3rdPressTop4Goal = row3rdPressTop.Cells[12];
            cell3rdPressTop4Goal.Format.Font.Size = 6.5;
            cell3rdPressTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressTop4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 22);
            cell3rdPressTop4Goal.AddParagraph($"{thirdPressTop4Goal.Goal1}");

            var cell3rdPressTop4Info = row3rdPressTop2.Cells[10];
            cell3rdPressTop4Info.Format.Font.Size = 6.5;
            cell3rdPressTop4Info.MergeRight = 2;
            cell3rdPressTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop4Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressTop4 != null)
            {
                cell3rdPressTop4Info.AddParagraph($"{clothing3rdPressTop4.Manufacturer} {clothing3rdPressTop4.Serial_Number}");
            }
            else cell3rdPressTop4Info.AddParagraph("NA");
            #endregion

            #endregion // 3rd Press Top Position

            #region 3rd Press Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdPressBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 1);
            var clothing3rdPressBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 2);
            var clothing3rdPressBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 3);
            var clothing3rdPressBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 4);
            var row3rdPressBottom = weeklyPMTable.AddRow();
            row3rdPressBottom.Height = 10;
            var row3rdPressBottom2 = weeklyPMTable.AddRow();
            row3rdPressBottom2.Height = 10;

            var cell3rdPressBottom = row3rdPressBottom.Cells[0];
            cell3rdPressBottom.Format.Font.Size = 6.5;
            cell3rdPressBottom.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom.MergeDown = 1;
            cell3rdPressBottom.AddParagraph("3RD PRESS BOTTOM");

            #region 3rd Press Bottom PM 1
            // ----- 3rd Press Bottom Position PM 1 ----- //
            var cell3rdPressBottom1Past = row3rdPressBottom.Cells[1];
            cell3rdPressBottom1Past.Format.Font.Size = 6.5;
            cell3rdPressBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 23 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressBottom1Past = thirdPressBottom1Past != null ? Math.Round((double)thirdPressBottom1Past) : 0;
            cell3rdPressBottom1Past.AddParagraph($"{thirdPressBottom1Past}");

            var cell3rdPressBottom1Current = row3rdPressBottom.Cells[2];
            cell3rdPressBottom1Current.Format.Font.Size = 6.5;
            cell3rdPressBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom1Age = 0;
            if (clothing3rdPressBottom1 != null)
            {
                thirdPressBottom1Age = clothing3rdPressBottom1.Age != null ? Convert.ToInt32(clothing3rdPressBottom1.Age) : 0;
            }
            cell3rdPressBottom1Current.AddParagraph($"{thirdPressBottom1Age}");

            var cell3rdPressBottom1Goal = row3rdPressBottom.Cells[3];
            cell3rdPressBottom1Goal.Format.Font.Size = 6.5;
            cell3rdPressBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottomlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 23);
            cell3rdPressBottom1Goal.AddParagraph($"{thirdPressBottomlGoal.Goal1}");

            var cell3rdPressBottom1Info = row3rdPressBottom2.Cells[1];
            cell3rdPressBottom1Info.Format.Font.Size = 6.5;
            cell3rdPressBottom1Info.MergeRight = 2;
            cell3rdPressBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom1Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressBottom1 != null)
            {
                cell3rdPressBottom1Info.AddParagraph($"{clothing3rdPressBottom1.Manufacturer} {clothing3rdPressBottom1.Serial_Number}");
            }
            else cell3rdPressBottom1Info.AddParagraph("NA");
            #endregion

            #region 3rd Press Bottom PM 2
            // ----- 3rd Press Bottom Position PM 2 ----- //
            var cell3rdPressBottom2Past = row3rdPressBottom.Cells[4];
            cell3rdPressBottom2Past.Format.Font.Size = 6.5;
            cell3rdPressBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 23 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressBottom2Past = thirdPressBottom2Past != null ? Math.Round((double)thirdPressBottom2Past) : 0;
            cell3rdPressBottom2Past.AddParagraph($"{thirdPressBottom2Past}");

            var cell3rdPressBottom2Current = row3rdPressBottom.Cells[5];
            cell3rdPressBottom2Current.Format.Font.Size = 6.5;
            cell3rdPressBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom2Age = 0;
            if (clothing3rdPressBottom2 != null)
            {
                thirdPressBottom2Age = clothing3rdPressBottom2.Age != null ? Convert.ToInt32(clothing3rdPressBottom2.Age) : 0;
            }
            cell3rdPressBottom2Current.AddParagraph($"{thirdPressBottom2Age}");

            var cell3rdPressBottom2Goal = row3rdPressBottom.Cells[6];
            cell3rdPressBottom2Goal.Format.Font.Size = 6.5;
            cell3rdPressBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 23);
            cell3rdPressBottom2Goal.AddParagraph($"{thirdPressBottom2Goal.Goal1}");

            var cell3rdPressBottom2Info = row3rdPressBottom2.Cells[4];
            cell3rdPressBottom2Info.Format.Font.Size = 6.5;
            cell3rdPressBottom2Info.MergeRight = 2;
            cell3rdPressBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom2Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressBottom2 != null)
            {
                cell3rdPressBottom2Info.AddParagraph($"{clothing3rdPressBottom2.Manufacturer} {clothing3rdPressBottom2.Serial_Number}");
            }
            else cell3rdPressBottom2Info.AddParagraph("NA");
            #endregion

            #region 3rd Press Bottom PM 3
            // ----- 3rd Press Bottom Position PM 3 ----- //
            var cell3rdPressBottom3Past = row3rdPressBottom.Cells[7];
            cell3rdPressBottom3Past.Format.Font.Size = 6.5;
            cell3rdPressBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 23 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressBottom3Past = thirdPressBottom3Past != null ? Math.Round((double)thirdPressBottom3Past) : 0;
            cell3rdPressBottom3Past.AddParagraph($"{thirdPressBottom3Past}");

            var cell3rdPressBottom3Current = row3rdPressBottom.Cells[8];
            cell3rdPressBottom3Current.Format.Font.Size = 6.5;
            cell3rdPressBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom3Age = 0;
            if (clothing3rdPressBottom3 != null)
            {
                thirdPressBottom3Age = clothing3rdPressBottom3.Age != null ? Convert.ToInt32(clothing3rdPressBottom3.Age) : 0;
            }
            cell3rdPressBottom3Current.AddParagraph($"{thirdPressBottom3Age}");

            var cell3rdPressBottom3Goal = row3rdPressBottom.Cells[9];
            cell3rdPressBottom3Goal.Format.Font.Size = 6.5;
            cell3rdPressBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 23);
            cell3rdPressBottom3Goal.AddParagraph($"{thirdPressBottom3Goal.Goal1}");

            var cell3rdPressBottom3Info = row3rdPressBottom2.Cells[7];
            cell3rdPressBottom3Info.Format.Font.Size = 6.5;
            cell3rdPressBottom3Info.MergeRight = 2;
            cell3rdPressBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom3Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressBottom3 != null)
            {
                cell3rdPressBottom3Info.AddParagraph($"{clothing3rdPressBottom3.Manufacturer} {clothing3rdPressBottom3.Serial_Number}");
            }
            else cell3rdPressBottom3Info.AddParagraph("NA");
            #endregion

            #region 3rd Press Bottom PM 4
            // ----- 3rd Press Bottom Position PM 4 ----- //
            var cell3rdPressBottom4Past = row3rdPressBottom.Cells[10];
            cell3rdPressBottom4Past.Format.Font.Size = 6.5;
            cell3rdPressBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 23 && c.StatusID == 3).Select(c => c.Age).Average();
            thirdPressBottom4Past = thirdPressBottom4Past != null ? Math.Round((double)thirdPressBottom4Past) : 0;
            cell3rdPressBottom4Past.AddParagraph($"{thirdPressBottom4Past}");

            var cell3rdPressBottom4Current = row3rdPressBottom.Cells[11];
            cell3rdPressBottom4Current.Format.Font.Size = 6.5;
            cell3rdPressBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom4Age = 0;
            if (clothing3rdPressBottom4 != null)
            {
                thirdPressBottom4Age = clothing3rdPressBottom4.Age != null ? Convert.ToInt32(clothing3rdPressBottom4.Age) : 0;
            }
            cell3rdPressBottom4Current.AddParagraph($"{thirdPressBottom4Age}");

            var cell3rdPressBottom4Goal = row3rdPressBottom.Cells[12];
            cell3rdPressBottom4Goal.Format.Font.Size = 6.5;
            cell3rdPressBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            var thirdPressBottom4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 23);
            cell3rdPressBottom4Goal.AddParagraph($"{thirdPressBottom4Goal.Goal1}");

            var cell3rdPressBottom4Info = row3rdPressBottom2.Cells[10];
            cell3rdPressBottom4Info.Format.Font.Size = 6.5;
            cell3rdPressBottom4Info.MergeRight = 2;
            cell3rdPressBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom4Info.Shading.Color = Colors.LightBlue;
            if (clothing3rdPressBottom4 != null)
            {
                cell3rdPressBottom4Info.AddParagraph($"{clothing3rdPressBottom4.Manufacturer} {clothing3rdPressBottom4.Serial_Number}");
            }
            else cell3rdPressBottom4Info.AddParagraph("NA");
            #endregion

            #endregion // 3rd Press Bottom Position

            #region Smooth Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSmoothTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 1);
            var clothingSmoothTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 2);
            var clothingSmoothTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 3);
            var clothingSmoothTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 4);
            var rowSmoothTop = weeklyPMTable.AddRow();
            rowSmoothTop.Height = 10;
            var rowSmoothTop2 = weeklyPMTable.AddRow();
            rowSmoothTop2.Height = 10;

            var cellSmoothTop = rowSmoothTop.Cells[0];
            cellSmoothTop.Format.Font.Size = 6.5;
            cellSmoothTop.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop.MergeDown = 1;
            cellSmoothTop.AddParagraph("SMOOTH TOP");

            #region Smooth Top PM 1
            // ----- Smooth Top Position PM 1 ----- //
            var cellSmoothTop1Past = rowSmoothTop.Cells[1];
            cellSmoothTop1Past.Format.Font.Size = 6.5;
            cellSmoothTop1Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 24 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothTop1Past = smoothTop1Past != null ? Math.Round((double)smoothTop1Past) : 0;
            cellSmoothTop1Past.AddParagraph($"{smoothTop1Past}");

            var cellSmoothTop1Current = rowSmoothTop.Cells[2];
            cellSmoothTop1Current.Format.Font.Size = 6.5;
            cellSmoothTop1Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop1Age = 0;
            if (clothingSmoothTop1 != null)
            {
                smoothTop1Age = clothingSmoothTop1.Age != null ? Convert.ToInt32(clothingSmoothTop1.Age) : 0;
            }
            cellSmoothTop1Current.AddParagraph($"{smoothTop1Age}");

            var cellSmoothTop1Goal = rowSmoothTop.Cells[3];
            cellSmoothTop1Goal.Format.Font.Size = 6.5;
            cellSmoothTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothToplGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 24);
            cellSmoothTop1Goal.AddParagraph($"{smoothToplGoal.Goal1}");

            var cellSmoothTop1Info = rowSmoothTop2.Cells[1];
            cellSmoothTop1Info.Format.Font.Size = 6.5;
            cellSmoothTop1Info.MergeRight = 2;
            cellSmoothTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop1Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothTop1 != null)
            {
                cellSmoothTop1Info.AddParagraph($"{clothingSmoothTop1.Manufacturer} {clothingSmoothTop1.Serial_Number}");
            }
            else cellSmoothTop1Info.AddParagraph("NA");
            #endregion

            #region Smooth Top PM 2
            // ----- Smooth Top Position PM 2 ----- //
            var cellSmoothTop2Past = rowSmoothTop.Cells[4];
            cellSmoothTop2Past.Format.Font.Size = 6.5;
            cellSmoothTop2Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 24 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothTop2Past = smoothTop2Past != null ? Math.Round((double)smoothTop2Past) : 0;
            cellSmoothTop2Past.AddParagraph($"{smoothTop2Past}");

            var cellSmoothTop2Current = rowSmoothTop.Cells[5];
            cellSmoothTop2Current.Format.Font.Size = 6.5;
            cellSmoothTop2Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop2Age = 0;
            if (clothingSmoothTop2 != null)
            {
                smoothTop2Age = clothingSmoothTop2.Age != null ? Convert.ToInt32(clothingSmoothTop2.Age) : 0;
            }
            cellSmoothTop2Current.AddParagraph($"{smoothTop2Age}");

            var cellSmoothTop2Goal = rowSmoothTop.Cells[6];
            cellSmoothTop2Goal.Format.Font.Size = 6.5;
            cellSmoothTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 24);
            cellSmoothTop2Goal.AddParagraph($"{smoothTop2Goal.Goal1}");

            var cellSmoothTop2Info = rowSmoothTop2.Cells[4];
            cellSmoothTop2Info.Format.Font.Size = 6.5;
            cellSmoothTop2Info.MergeRight = 2;
            cellSmoothTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop2Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothTop2 != null)
            {
                cellSmoothTop2Info.AddParagraph($"{clothingSmoothTop2.Manufacturer} {clothingSmoothTop2.Serial_Number}");
            }
            else cellSmoothTop2Info.AddParagraph("NA");
            #endregion

            #region Smooth Top PM 3
            // ----- Smooth Top Position PM 3 ----- //
            var cellSmoothTop3Past = rowSmoothTop.Cells[7];
            cellSmoothTop3Past.Format.Font.Size = 6.5;
            cellSmoothTop3Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 24 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothTop3Past = smoothTop3Past != null ? Math.Round((double)smoothTop3Past) : 0;
            cellSmoothTop3Past.AddParagraph($"{smoothTop3Past}");

            var cellSmoothTop3Current = rowSmoothTop.Cells[8];
            cellSmoothTop3Current.Format.Font.Size = 6.5;
            cellSmoothTop3Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop3Age = 0;
            if (clothingSmoothTop3 != null)
            {
                smoothTop3Age = clothingSmoothTop3.Age != null ? Convert.ToInt32(clothingSmoothTop3.Age) : 0;
            }
            cellSmoothTop3Current.AddParagraph($"{smoothTop3Age}");

            var cellSmoothTop3Goal = rowSmoothTop.Cells[9];
            cellSmoothTop3Goal.Format.Font.Size = 6.5;
            cellSmoothTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 24);
            cellSmoothTop3Goal.AddParagraph($"{smoothTop3Goal.Goal1}");

            var cellSmoothTop3Info = rowSmoothTop2.Cells[7];
            cellSmoothTop3Info.Format.Font.Size = 6.5;
            cellSmoothTop3Info.MergeRight = 2;
            cellSmoothTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop3Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothTop3 != null)
            {
                cellSmoothTop3Info.AddParagraph($"{clothingSmoothTop3.Manufacturer} {clothingSmoothTop3.Serial_Number}");
            }
            else cellSmoothTop3Info.AddParagraph("NA");
            #endregion

            #region Smooth Top PM 4
            // ----- Smooth Top Position PM 4 ----- //
            var cellSmoothTop4Past = rowSmoothTop.Cells[10];
            cellSmoothTop4Past.Format.Font.Size = 6.5;
            cellSmoothTop4Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 24 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothTop4Past = smoothTop4Past != null ? Math.Round((double)smoothTop4Past) : 0;
            cellSmoothTop4Past.AddParagraph($"{smoothTop3Past}");

            var cellSmoothTop4Current = rowSmoothTop.Cells[11];
            cellSmoothTop4Current.Format.Font.Size = 6.5;
            cellSmoothTop4Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop4Age = 0;
            if (clothingSmoothTop4 != null)
            {
                smoothTop4Age = clothingSmoothTop4.Age != null ? Convert.ToInt32(clothingSmoothTop4.Age) : 0;
            }
            cellSmoothTop4Current.AddParagraph($"{smoothTop4Age}");

            var cellSmoothTop4Goal = rowSmoothTop.Cells[12];
            cellSmoothTop4Goal.Format.Font.Size = 6.5;
            cellSmoothTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothTop4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 24);
            cellSmoothTop4Goal.AddParagraph($"{smoothTop4Goal.Goal1}");

            var cellSmoothTop4Info = rowSmoothTop2.Cells[10];
            cellSmoothTop4Info.Format.Font.Size = 6.5;
            cellSmoothTop4Info.MergeRight = 2;
            cellSmoothTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop4Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothTop4 != null)
            {
                cellSmoothTop4Info.AddParagraph($"{clothingSmoothTop4.Manufacturer} {clothingSmoothTop4.Serial_Number}");
            }
            else cellSmoothTop4Info.AddParagraph("NA");
            #endregion

            #endregion // Smooth Top Position

            #region Smooth Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSmoothBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 1);
            var clothingSmoothBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 2);
            var clothingSmoothBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 3);
            var clothingSmoothBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 4);
            var rowSmoothBottom = weeklyPMTable.AddRow();
            rowSmoothBottom.Height = 10;
            var rowSmoothBottom2 = weeklyPMTable.AddRow();
            rowSmoothBottom2.Height = 10;

            var cellSmoothBottom = rowSmoothBottom.Cells[0];
            cellSmoothBottom.Format.Font.Size = 6.5;
            cellSmoothBottom.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom.MergeDown = 1;
            cellSmoothBottom.AddParagraph("SMOOTH BOTTOM");

            #region Smooth Bottom PM 1
            // ----- Smooth Bottom Position PM 1 ----- //
            var cellSmoothBottom1Past = rowSmoothBottom.Cells[1];
            cellSmoothBottom1Past.Format.Font.Size = 6.5;
            cellSmoothBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 25 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothBottom1Past = smoothBottom1Past != null ? Math.Round((double)smoothBottom1Past) : 0;
            cellSmoothBottom1Past.AddParagraph($"{smoothBottom1Past}");

            var cellSmoothBottom1Current = rowSmoothBottom.Cells[2];
            cellSmoothBottom1Current.Format.Font.Size = 6.5;
            cellSmoothBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom1Age = 0;
            if (clothingSmoothBottom1 != null)
            {
                smoothBottom1Age = clothingSmoothBottom1.Age != null ? Convert.ToInt32(clothingSmoothBottom1.Age) : 0;
            }
            cellSmoothBottom1Current.AddParagraph($"{smoothBottom1Age}");

            var cellSmoothBottom1Goal = rowSmoothBottom.Cells[3];
            cellSmoothBottom1Goal.Format.Font.Size = 6.5;
            cellSmoothBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottomlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 25);
            cellSmoothBottom1Goal.AddParagraph($"{smoothBottomlGoal.Goal1}");

            var cellSmoothBottom1Info = rowSmoothBottom2.Cells[1];
            cellSmoothBottom1Info.Format.Font.Size = 6.5;
            cellSmoothBottom1Info.MergeRight = 2;
            cellSmoothBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom1Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothBottom1 != null)
            {
                cellSmoothBottom1Info.AddParagraph($"{clothingSmoothBottom1.Manufacturer} {clothingSmoothBottom1.Serial_Number}");
            }
            else cellSmoothBottom1Info.AddParagraph("NA");
            #endregion

            #region Smooth Bottom PM 2
            // ----- Smooth Bottom Position PM 2 ----- //
            var cellSmoothBottom2Past = rowSmoothBottom.Cells[4];
            cellSmoothBottom2Past.Format.Font.Size = 6.5;
            cellSmoothBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 25 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothBottom2Past = smoothBottom2Past != null ? Math.Round((double)smoothBottom2Past) : 0;
            cellSmoothBottom2Past.AddParagraph($"{smoothBottom2Past}");

            var cellSmoothBottom2Current = rowSmoothBottom.Cells[5];
            cellSmoothBottom2Current.Format.Font.Size = 6.5;
            cellSmoothBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom2Age = 0;
            if (clothingSmoothBottom2 != null)
            {
                smoothBottom2Age = clothingSmoothBottom2.Age != null ? Convert.ToInt32(clothingSmoothBottom2.Age) : 0;
            }
            cellSmoothBottom2Current.AddParagraph($"{smoothBottom2Age}");

            var cellSmoothBottom2Goal = rowSmoothBottom.Cells[6];
            cellSmoothBottom2Goal.Format.Font.Size = 6.5;
            cellSmoothBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 25);
            cellSmoothBottom2Goal.AddParagraph($"{smoothBottom2Goal.Goal1}");

            var cellSmoothBottom2Info = rowSmoothBottom2.Cells[4];
            cellSmoothBottom2Info.Format.Font.Size = 6.5;
            cellSmoothBottom2Info.MergeRight = 2;
            cellSmoothBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom2Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothBottom2 != null)
            {
                cellSmoothBottom2Info.AddParagraph($"{clothingSmoothBottom2.Manufacturer} {clothingSmoothBottom2.Serial_Number}");
            }
            else cellSmoothBottom2Info.AddParagraph("NA");
            #endregion

            #region Smooth Bottom PM 3
            // ----- Smooth Bottom Position PM 3 ----- //
            var cellSmoothBottom3Past = rowSmoothBottom.Cells[7];
            cellSmoothBottom3Past.Format.Font.Size = 6.5;
            cellSmoothBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 25 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothBottom3Past = smoothBottom3Past != null ? Math.Round((double)smoothBottom3Past) : 0;
            cellSmoothBottom3Past.AddParagraph($"{smoothBottom3Past}");

            var cellSmoothBottom3Current = rowSmoothBottom.Cells[8];
            cellSmoothBottom3Current.Format.Font.Size = 6.5;
            cellSmoothBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom3Age = 0;
            if (clothingSmoothBottom3 != null)
            {
                smoothBottom3Age = clothingSmoothBottom3.Age != null ? Convert.ToInt32(clothingSmoothBottom3.Age) : 0;
            }
            cellSmoothBottom3Current.AddParagraph($"{smoothBottom3Age}");

            var cellSmoothBottom3Goal = rowSmoothBottom.Cells[9];
            cellSmoothBottom3Goal.Format.Font.Size = 6.5;
            cellSmoothBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 25);
            cellSmoothBottom3Goal.AddParagraph($"{smoothBottom3Goal.Goal1}");

            var cellSmoothBottom3Info = rowSmoothBottom2.Cells[7];
            cellSmoothBottom3Info.Format.Font.Size = 6.5;
            cellSmoothBottom3Info.MergeRight = 2;
            cellSmoothBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom3Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothBottom3 != null)
            {
                cellSmoothBottom3Info.AddParagraph($"{clothingSmoothBottom3.Manufacturer} {clothingSmoothBottom3.Serial_Number}");
            }
            else cellSmoothBottom3Info.AddParagraph("NA");
            #endregion

            #region Smooth Bottom PM 4
            // ----- Smooth Bottom Position PM 4 ----- //
            var cellSmoothBottom4Past = rowSmoothBottom.Cells[10];
            cellSmoothBottom4Past.Format.Font.Size = 6.5;
            cellSmoothBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 25 && c.StatusID == 3).Select(c => c.Age).Average();
            smoothBottom4Past = smoothBottom4Past != null ? Math.Round((double)smoothBottom4Past) : 0;
            cellSmoothBottom4Past.AddParagraph($"{smoothBottom4Past}");

            var cellSmoothBottom4Current = rowSmoothBottom.Cells[11];
            cellSmoothBottom4Current.Format.Font.Size = 6.5;
            cellSmoothBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom4Age = 0;
            if (clothingSmoothBottom4 != null)
            {
                smoothBottom4Age = clothingSmoothBottom4.Age != null ? Convert.ToInt32(clothingSmoothBottom4.Age) : 0;
            }
            cellSmoothBottom4Current.AddParagraph($"{smoothBottom4Age}");

            var cellSmoothBottom4Goal = rowSmoothBottom.Cells[12];
            cellSmoothBottom4Goal.Format.Font.Size = 6.5;
            cellSmoothBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            var smoothBottom4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 25);
            cellSmoothBottom4Goal.AddParagraph($"{smoothBottom4Goal.Goal1}");

            var cellSmoothBottom4Info = rowSmoothBottom2.Cells[10];
            cellSmoothBottom4Info.Format.Font.Size = 6.5;
            cellSmoothBottom4Info.MergeRight = 2;
            cellSmoothBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom4Info.Shading.Color = Colors.LightBlue;
            if (clothingSmoothBottom4 != null)
            {
                cellSmoothBottom4Info.AddParagraph($"{clothingSmoothBottom4.Manufacturer} {clothingSmoothBottom4.Serial_Number}");
            }
            else cellSmoothBottom4Info.AddParagraph("NA");
            #endregion

            #endregion // Smooth Bottom Position

            #region Hard Size Press Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingHardSizePress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 1);
            var clothingHardSizePress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 2);
            var clothingHardSizePress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 3);
            var clothingHardSizePress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 4);
            var rowHardSizePress = weeklyPMTable.AddRow();
            rowHardSizePress.Height = 10;
            var rowHardSizePress2 = weeklyPMTable.AddRow();
            rowHardSizePress2.Height = 10;

            var cellHardSizePress = rowHardSizePress.Cells[0];
            cellHardSizePress.Format.Font.Size = 6.5;
            cellHardSizePress.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress.MergeDown = 1;
            cellHardSizePress.AddParagraph("HARD SIZE PRESS (TOP)");

            #region Hard Size Press PM 1
            // ----- Hard Size Press Position PM 1 ----- //
            var cellHardSizePress1Past = rowHardSizePress.Cells[1];
            cellHardSizePress1Past.Format.Font.Size = 6.5;
            cellHardSizePress1Past.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 26 && c.StatusID == 3).Select(c => c.Age).Average();
            hardSizePress1Past = hardSizePress1Past != null ? Math.Round((double)hardSizePress1Past) : 0;
            cellHardSizePress1Past.AddParagraph($"{hardSizePress1Past}");

            var cellHardSizePress1Current = rowHardSizePress.Cells[2];
            cellHardSizePress1Current.Format.Font.Size = 6.5;
            cellHardSizePress1Current.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress1Age = 0;
            if (clothingHardSizePress1 != null)
            {
                hardSizePress1Age = clothingHardSizePress1.Age != null ? Convert.ToInt32(clothingHardSizePress1.Age) : 0;
            }
            cellHardSizePress1Current.AddParagraph($"{hardSizePress1Age}");

            var cellHardSizePress1Goal = rowHardSizePress.Cells[3];
            cellHardSizePress1Goal.Format.Font.Size = 6.5;
            cellHardSizePress1Goal.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePresslGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 26);
            cellHardSizePress1Goal.AddParagraph($"{hardSizePresslGoal.Goal1}");

            var cellHardSizePress1Info = rowHardSizePress2.Cells[1];
            cellHardSizePress1Info.Format.Font.Size = 6.5;
            cellHardSizePress1Info.MergeRight = 2;
            cellHardSizePress1Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress1Info.Shading.Color = Colors.LightBlue;
            if (clothingHardSizePress1 != null)
            {
                cellHardSizePress1Info.AddParagraph($"{clothingHardSizePress1.Manufacturer} {clothingHardSizePress1.Serial_Number}");
            }
            else cellHardSizePress1Info.AddParagraph("NA");
            #endregion

            #region Hard Size Press PM 2
            // ----- Hard Size Press Position PM 2 ----- //
            var cellHardSizePress2Past = rowHardSizePress.Cells[4];
            cellHardSizePress2Past.Format.Font.Size = 6.5;
            cellHardSizePress2Past.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 26 && c.StatusID == 3).Select(c => c.Age).Average();
            hardSizePress2Past = hardSizePress2Past != null ? Math.Round((double)hardSizePress2Past) : 0;
            cellHardSizePress2Past.AddParagraph($"{hardSizePress2Past}");

            var cellHardSizePress2Current = rowHardSizePress.Cells[5];
            cellHardSizePress2Current.Format.Font.Size = 6.5;
            cellHardSizePress2Current.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress2Age = 0;
            if (clothingHardSizePress2 != null)
            {
                hardSizePress2Age = clothingHardSizePress2.Age != null ? Convert.ToInt32(clothingHardSizePress2.Age) : 0;
            }
            cellHardSizePress2Current.AddParagraph($"{hardSizePress2Age}");

            var cellHardSizePress2Goal = rowHardSizePress.Cells[6];
            cellHardSizePress2Goal.Format.Font.Size = 6.5;
            cellHardSizePress2Goal.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 26);
            cellHardSizePress2Goal.AddParagraph($"{hardSizePress2Goal.Goal1}");

            var cellHardSizePress2Info = rowHardSizePress2.Cells[4];
            cellHardSizePress2Info.Format.Font.Size = 6.5;
            cellHardSizePress2Info.MergeRight = 2;
            cellHardSizePress2Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress2Info.Shading.Color = Colors.LightBlue;
            if (clothingHardSizePress2 != null)
            {
                cellHardSizePress2Info.AddParagraph($"{clothingHardSizePress2.Manufacturer} {clothingHardSizePress2.Serial_Number}");
            }
            else cellHardSizePress2Info.AddParagraph("NA");
            #endregion

            #region Hard Size Press PM 3
            // ----- Hard Size Press Position PM 3 ----- //
            var cellHardSizePress3Past = rowHardSizePress.Cells[7];
            cellHardSizePress3Past.Format.Font.Size = 6.5;
            cellHardSizePress3Past.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 26 && c.StatusID == 3).Select(c => c.Age).Average();
            hardSizePress3Past = hardSizePress3Past != null ? Math.Round((double)hardSizePress3Past) : 0;
            cellHardSizePress3Past.AddParagraph($"{hardSizePress3Past}");

            var cellHardSizePress3Current = rowHardSizePress.Cells[8];
            cellHardSizePress3Current.Format.Font.Size = 6.5;
            cellHardSizePress3Current.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress3Age = 0;
            if (clothingHardSizePress3 != null)
            {
                hardSizePress3Age = clothingHardSizePress3.Age != null ? Convert.ToInt32(clothingHardSizePress3.Age) : 0;
            }
            cellHardSizePress3Current.AddParagraph($"{hardSizePress3Age}");

            var cellHardSizePress3Goal = rowHardSizePress.Cells[9];
            cellHardSizePress3Goal.Format.Font.Size = 6.5;
            cellHardSizePress3Goal.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 26);
            cellHardSizePress3Goal.AddParagraph($"{hardSizePress3Goal.Goal1}");

            var cellHardSizePress3Info = rowHardSizePress2.Cells[7];
            cellHardSizePress3Info.Format.Font.Size = 6.5;
            cellHardSizePress3Info.MergeRight = 2;
            cellHardSizePress3Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress3Info.Shading.Color = Colors.LightBlue;
            if (clothingHardSizePress3 != null)
            {
                cellHardSizePress3Info.AddParagraph($"{clothingHardSizePress3.Manufacturer} {clothingHardSizePress3.Serial_Number}");
            }
            else cellHardSizePress3Info.AddParagraph("NA");
            #endregion

            #region Hard Size Press PM 4
            // ----- Hard Size Press Position PM 4 ----- //
            var cellHardSizePress4Past = rowHardSizePress.Cells[10];
            cellHardSizePress4Past.Format.Font.Size = 6.5;
            cellHardSizePress4Past.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 26 && c.StatusID == 3).Select(c => c.Age).Average();
            hardSizePress4Past = hardSizePress4Past != null ? Math.Round((double)hardSizePress4Past) : 0;
            cellHardSizePress4Past.AddParagraph($"{hardSizePress4Past}");

            var cellHardSizePress4Current = rowHardSizePress.Cells[11];
            cellHardSizePress4Current.Format.Font.Size = 6.5;
            cellHardSizePress4Current.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress4Age = 0;
            if (clothingHardSizePress4 != null)
            {
                hardSizePress4Age = clothingHardSizePress4.Age != null ? Convert.ToInt32(clothingHardSizePress4.Age) : 0;
            }
            cellHardSizePress4Current.AddParagraph($"{hardSizePress4Age}");

            var cellHardSizePress4Goal = rowHardSizePress.Cells[12];
            cellHardSizePress4Goal.Format.Font.Size = 6.5;
            cellHardSizePress4Goal.VerticalAlignment = VerticalAlignment.Center;
            var hardSizePress4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 26);
            cellHardSizePress4Goal.AddParagraph($"{hardSizePress4Goal.Goal1}");

            var cellHardSizePress4Info = rowHardSizePress2.Cells[10];
            cellHardSizePress4Info.Format.Font.Size = 6.5;
            cellHardSizePress4Info.MergeRight = 2;
            cellHardSizePress4Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress4Info.Shading.Color = Colors.LightBlue;
            if (clothingHardSizePress4 != null)
            {
                cellHardSizePress4Info.AddParagraph($"{clothingHardSizePress4.Manufacturer} {clothingHardSizePress4.Serial_Number}");
            }
            else cellHardSizePress4Info.AddParagraph("NA");
            #endregion

            #endregion // Hard Size Press Position

            #region Soft Size Press Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSoftSizePress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 1);
            var clothingSoftSizePress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 2);
            var clothingSoftSizePress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 3);
            var clothingSoftSizePress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 4);
            var rowSoftSizePress = weeklyPMTable.AddRow();
            rowSoftSizePress.Height = 10;
            var rowSoftSizePress2 = weeklyPMTable.AddRow();
            rowSoftSizePress2.Height = 10;

            var cellSoftSizePress = rowSoftSizePress.Cells[0];
            cellSoftSizePress.Format.Font.Size = 6.5;
            cellSoftSizePress.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress.MergeDown = 1;
            cellSoftSizePress.AddParagraph("SOFT SIZE PRESS (BOTTOM)");

            #region Soft Size Press PM 1
            // ----- Soft Size Press Position PM 1 ----- //
            var cellSoftSizePress1Past = rowSoftSizePress.Cells[1];
            cellSoftSizePress1Past.Format.Font.Size = 6.5;
            cellSoftSizePress1Past.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 27 && c.StatusID == 3).Select(c => c.Age).Average();
            softSizePress1Past = softSizePress1Past != null ? Math.Round((double)softSizePress1Past) : 0;
            cellSoftSizePress1Past.AddParagraph($"{softSizePress1Past}");

            var cellSoftSizePress1Current = rowSoftSizePress.Cells[2];
            cellSoftSizePress1Current.Format.Font.Size = 6.5;
            cellSoftSizePress1Current.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress1Age = 0;
            if (clothingSoftSizePress1 != null)
            {
                softSizePress1Age = clothingSoftSizePress1.Age != null ? Convert.ToInt32(clothingSoftSizePress1.Age) : 0;
            }
            cellSoftSizePress1Current.AddParagraph($"{softSizePress1Age}");

            var cellSoftSizePress1Goal = rowSoftSizePress.Cells[3];
            cellSoftSizePress1Goal.Format.Font.Size = 6.5;
            cellSoftSizePress1Goal.VerticalAlignment = VerticalAlignment.Center;
            var softSizePresslGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 27);
            cellSoftSizePress1Goal.AddParagraph($"{softSizePresslGoal.Goal1}");

            var cellSoftSizePress1Info = rowSoftSizePress2.Cells[1];
            cellSoftSizePress1Info.Format.Font.Size = 6.5;
            cellSoftSizePress1Info.MergeRight = 2;
            cellSoftSizePress1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress1Info.Shading.Color = Colors.LightBlue;
            if (clothingSoftSizePress1 != null)
            {
                cellSoftSizePress1Info.AddParagraph($"{clothingSoftSizePress1.Manufacturer} {clothingSoftSizePress1.Serial_Number}");
            }
            else cellSoftSizePress1Info.AddParagraph("NA");
            #endregion

            #region Soft Size Press PM 2
            // ----- Soft Size Press Position PM 2 ----- //
            var cellSoftSizePress2Past = rowSoftSizePress.Cells[4];
            cellSoftSizePress2Past.Format.Font.Size = 6.5;
            cellSoftSizePress2Past.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 27 && c.StatusID == 3).Select(c => c.Age).Average();
            softSizePress2Past = softSizePress2Past != null ? Math.Round((double)softSizePress2Past) : 0;
            cellSoftSizePress2Past.AddParagraph($"{softSizePress2Past}");

            var cellSoftSizePress2Current = rowSoftSizePress.Cells[5];
            cellSoftSizePress2Current.Format.Font.Size = 6.5;
            cellSoftSizePress2Current.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress2Age = 0;
            if (clothingSoftSizePress2 != null)
            {
                softSizePress2Age = clothingSoftSizePress2.Age != null ? Convert.ToInt32(clothingSoftSizePress2.Age) : 0;
            }
            cellSoftSizePress2Current.AddParagraph($"{softSizePress2Age}");

            var cellSoftSizePress2Goal = rowSoftSizePress.Cells[6];
            cellSoftSizePress2Goal.Format.Font.Size = 6.5;
            cellSoftSizePress2Goal.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 27);
            cellSoftSizePress2Goal.AddParagraph($"{softSizePress2Goal.Goal1}");

            var cellSoftSizePress2Info = rowSoftSizePress2.Cells[4];
            cellSoftSizePress2Info.Format.Font.Size = 6.5;
            cellSoftSizePress2Info.MergeRight = 2;
            cellSoftSizePress2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress2Info.Shading.Color = Colors.LightBlue;
            if (clothingSoftSizePress2 != null)
            {
                cellSoftSizePress2Info.AddParagraph($"{clothingSoftSizePress2.Manufacturer} {clothingSoftSizePress2.Serial_Number}");
            }
            else cellSoftSizePress2Info.AddParagraph("NA");
            #endregion

            #region Soft Size Press PM 3
            // ----- Soft Size Press Position PM 3 ----- //
            var cellSoftSizePress3Past = rowSoftSizePress.Cells[7];
            cellSoftSizePress3Past.Format.Font.Size = 6.5;
            cellSoftSizePress3Past.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 27 && c.StatusID == 3).Select(c => c.Age).Average();
            softSizePress3Past = softSizePress3Past != null ? Math.Round((double)softSizePress3Past) : 0;
            cellSoftSizePress3Past.AddParagraph($"{softSizePress3Past}");

            var cellSoftSizePress3Current = rowSoftSizePress.Cells[8];
            cellSoftSizePress3Current.Format.Font.Size = 6.5;
            cellSoftSizePress3Current.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress3Age = 0;
            if (clothingSoftSizePress3 != null)
            {
                softSizePress3Age = clothingSoftSizePress3.Age != null ? Convert.ToInt32(clothingSoftSizePress3.Age) : 0;
            }
            cellSoftSizePress3Current.AddParagraph($"{softSizePress3Age}");

            var cellSoftSizePress3Goal = rowSoftSizePress.Cells[9];
            cellSoftSizePress3Goal.Format.Font.Size = 6.5;
            cellSoftSizePress3Goal.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 27);
            cellSoftSizePress3Goal.AddParagraph($"{softSizePress3Goal.Goal1}");

            var cellSoftSizePress3Info = rowSoftSizePress2.Cells[7];
            cellSoftSizePress3Info.Format.Font.Size = 6.5;
            cellSoftSizePress3Info.MergeRight = 2;
            cellSoftSizePress3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress3Info.Shading.Color = Colors.LightBlue;
            if (clothingSoftSizePress3 != null)
            {
                cellSoftSizePress3Info.AddParagraph($"{clothingSoftSizePress3.Manufacturer} {clothingSoftSizePress3.Serial_Number}");
            }
            else cellSoftSizePress3Info.AddParagraph("NA");
            #endregion

            #region Soft Size Press PM 4
            // ----- Soft Size Press Position PM 4 ----- //
            var cellSoftSizePress4Past = rowSoftSizePress.Cells[10];
            cellSoftSizePress4Past.Format.Font.Size = 6.5;
            cellSoftSizePress4Past.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 27 && c.StatusID == 3).Select(c => c.Age).Average();
            softSizePress4Past = softSizePress4Past != null ? Math.Round((double)softSizePress4Past) : 0;
            cellSoftSizePress4Past.AddParagraph($"{softSizePress4Past}");

            var cellSoftSizePress4Current = rowSoftSizePress.Cells[11];
            cellSoftSizePress4Current.Format.Font.Size = 6.5;
            cellSoftSizePress4Current.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress4Age = 0;
            if (clothingSoftSizePress4 != null)
            {
                softSizePress4Age = clothingSoftSizePress4.Age != null ? Convert.ToInt32(clothingSoftSizePress4.Age) : 0;
            }
            cellSoftSizePress4Current.AddParagraph($"{softSizePress4Age}");

            var cellSoftSizePress4Goal = rowSoftSizePress.Cells[12];
            cellSoftSizePress4Goal.Format.Font.Size = 6.5;
            cellSoftSizePress4Goal.VerticalAlignment = VerticalAlignment.Center;
            var softSizePress4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 27);
            cellSoftSizePress4Goal.AddParagraph($"{softSizePress4Goal.Goal1}");

            var cellSoftSizePress4Info = rowSoftSizePress2.Cells[10];
            cellSoftSizePress4Info.Format.Font.Size = 6.5;
            cellSoftSizePress4Info.MergeRight = 2;
            cellSoftSizePress4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress4Info.Shading.Color = Colors.LightBlue;
            if (clothingSoftSizePress4 != null)
            {
                cellSoftSizePress4Info.AddParagraph($"{clothingSoftSizePress4.Manufacturer} {clothingSoftSizePress4.Serial_Number}");
            }
            else cellSoftSizePress4Info.AddParagraph("NA");
            #endregion

            #endregion // Soft Size Press Position

            #region Aquitherm Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingAquithermTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 1);
            var clothingAquithermTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 2);
            var clothingAquithermTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 3);
            var clothingAquithermTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 4);
            var rowAquithermTop = weeklyPMTable.AddRow();
            rowAquithermTop.Height = 10;
            var rowAquithermTop2 = weeklyPMTable.AddRow();
            rowAquithermTop2.Height = 10;

            var cellAquithermTop = rowAquithermTop.Cells[0];
            cellAquithermTop.Format.Font.Size = 6.5;
            cellAquithermTop.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop.MergeDown = 1;
            cellAquithermTop.AddParagraph("AQUATHERM TOP");

            #region Aquitherm Top PM 1
            // ----- Aquitherm Top Position PM 1 ----- //
            var cellAquithermTop1Past = rowAquithermTop.Cells[1];
            cellAquithermTop1Past.Format.Font.Size = 6.5;
            cellAquithermTop1Past.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 28 && c.StatusID == 3).Select(c => c.Age).Average();
            aquithermTop1Past = aquithermTop1Past != null ? Math.Round((double)aquithermTop1Past) : 0;
            cellAquithermTop1Past.AddParagraph($"{aquithermTop1Past}");

            var cellAquithermTop1Current = rowAquithermTop.Cells[2];
            cellAquithermTop1Current.Format.Font.Size = 6.5;
            cellAquithermTop1Current.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop1Age = 0;
            if (clothingAquithermTop1 != null)
            {
                aquithermTop1Age = clothingAquithermTop1.Age != null ? Convert.ToInt32(clothingAquithermTop1.Age) : 0;
            }
            cellAquithermTop1Current.AddParagraph($"{aquithermTop1Age}");

            var cellAquithermTop1Goal = rowAquithermTop.Cells[3];
            cellAquithermTop1Goal.Format.Font.Size = 6.5;
            cellAquithermTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            var aquithermToplGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 28);
            cellAquithermTop1Goal.AddParagraph($"{aquithermToplGoal.Goal1}");

            var cellAquithermTop1Info = rowAquithermTop2.Cells[1];
            cellAquithermTop1Info.Format.Font.Size = 6.5;
            cellAquithermTop1Info.MergeRight = 2;
            cellAquithermTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop1Info.Shading.Color = Colors.LightBlue;
            if (clothingAquithermTop1 != null)
            {
                cellAquithermTop1Info.AddParagraph($"{clothingAquithermTop1.Manufacturer} {clothingAquithermTop1.Serial_Number}");
            }
            else cellAquithermTop1Info.AddParagraph("NA");
            #endregion

            #region Aquitherm Top PM 2
            // ----- Aquitherm Top Position PM 2 ----- //
            var cellAquithermTop2Past = rowAquithermTop.Cells[4];
            cellAquithermTop2Past.Format.Font.Size = 6.5;
            cellAquithermTop2Past.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 28 && c.StatusID == 3).Select(c => c.Age).Average();
            aquithermTop2Past = aquithermTop2Past != null ? Math.Round((double)aquithermTop2Past) : 0;
            cellAquithermTop2Past.AddParagraph($"{aquithermTop2Past}");

            var cellAquithermTop2Current = rowAquithermTop.Cells[5];
            cellAquithermTop2Current.Format.Font.Size = 6.5;
            cellAquithermTop2Current.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop2Age = 0;
            if (clothingAquithermTop2 != null)
            {
                aquithermTop2Age = clothingAquithermTop2.Age != null ? Convert.ToInt32(clothingAquithermTop2.Age) : 0;
            }
            cellAquithermTop2Current.AddParagraph($"{aquithermTop2Age}");

            var cellAquithermTop2Goal = rowAquithermTop.Cells[6];
            cellAquithermTop2Goal.Format.Font.Size = 6.5;
            cellAquithermTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 28);
            cellAquithermTop2Goal.AddParagraph($"{aquithermTop2Goal.Goal1}");

            var cellAquithermTop2Info = rowAquithermTop2.Cells[4];
            cellAquithermTop2Info.Format.Font.Size = 6.5;
            cellAquithermTop2Info.MergeRight = 2;
            cellAquithermTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop2Info.Shading.Color = Colors.LightBlue;
            if (clothingAquithermTop2 != null)
            {
                cellAquithermTop2Info.AddParagraph($"{clothingAquithermTop2.Manufacturer} {clothingAquithermTop2.Serial_Number}");
            }
            else cellAquithermTop2Info.AddParagraph("NA");
            #endregion

            #region Aquitherm Top PM 3
            // ----- Aquitherm Top Position PM 3 ----- //
            var cellAquithermTop3Past = rowAquithermTop.Cells[7];
            cellAquithermTop3Past.Format.Font.Size = 6.5;
            cellAquithermTop3Past.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 28 && c.StatusID == 3).Select(c => c.Age).Average();
            aquithermTop3Past = aquithermTop3Past != null ? Math.Round((double)aquithermTop3Past) : 0;
            cellAquithermTop3Past.AddParagraph($"{aquithermTop3Past}");

            var cellAquithermTop3Current = rowAquithermTop.Cells[8];
            cellAquithermTop3Current.Format.Font.Size = 6.5;
            cellAquithermTop3Current.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop3Age = 0;
            if (clothingAquithermTop3 != null)
            {
                aquithermTop3Age = clothingAquithermTop3.Age != null ? Convert.ToInt32(clothingAquithermTop3.Age) : 0;
            }
            cellAquithermTop3Current.AddParagraph($"{aquithermTop3Age}");

            var cellAquithermTop3Goal = rowAquithermTop.Cells[9];
            cellAquithermTop3Goal.Format.Font.Size = 6.5;
            cellAquithermTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 28);
            cellAquithermTop3Goal.AddParagraph($"{aquithermTop3Goal.Goal1}");

            var cellAquithermTop3Info = rowAquithermTop2.Cells[7];
            cellAquithermTop3Info.Format.Font.Size = 6.5;
            cellAquithermTop3Info.MergeRight = 2;
            cellAquithermTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop3Info.Shading.Color = Colors.LightBlue;
            if (clothingAquithermTop3 != null)
            {
                cellAquithermTop3Info.AddParagraph($"{clothingAquithermTop3.Manufacturer} {clothingAquithermTop3.Serial_Number}");
            }
            else cellAquithermTop3Info.AddParagraph("NA");
            #endregion

            #region Aquitherm Top PM 4
            // ----- Aquitherm Top Position PM 4 ----- //
            var cellAquithermTop4Past = rowAquithermTop.Cells[10];
            cellAquithermTop4Past.Format.Font.Size = 6.5;
            cellAquithermTop4Past.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 28 && c.StatusID == 3).Select(c => c.Age).Average();
            aquithermTop4Past = aquithermTop4Past != null ? Math.Round((double)aquithermTop4Past) : 0;
            cellAquithermTop4Past.AddParagraph($"{aquithermTop3Past}");

            var cellAquithermTop4Current = rowAquithermTop.Cells[11];
            cellAquithermTop4Current.Format.Font.Size = 6.5;
            cellAquithermTop4Current.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop4Age = 0;
            if (clothingAquithermTop4 != null)
            {
                aquithermTop4Age = clothingAquithermTop4.Age != null ? Convert.ToInt32(clothingAquithermTop4.Age) : 0;
            }
            cellAquithermTop4Current.AddParagraph($"{aquithermTop4Age}");

            var cellAquithermTop4Goal = rowAquithermTop.Cells[12];
            cellAquithermTop4Goal.Format.Font.Size = 6.5;
            cellAquithermTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            var aquithermTop4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 28);
            cellAquithermTop4Goal.AddParagraph($"{aquithermTop4Goal.Goal1}");

            var cellAquithermTop4Info = rowAquithermTop2.Cells[10];
            cellAquithermTop4Info.Format.Font.Size = 6.5;
            cellAquithermTop4Info.MergeRight = 2;
            cellAquithermTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop4Info.Shading.Color = Colors.LightBlue;
            if (clothingAquithermTop4 != null)
            {
                cellAquithermTop4Info.AddParagraph($"{clothingAquithermTop4.Manufacturer} {clothingAquithermTop4.Serial_Number}");
            }
            else cellAquithermTop4Info.AddParagraph("NA");
            #endregion

            #endregion // Aquitherm Top Position

            #region Nibco Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingNibcoBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nibco roll" && c.PM_Number == 1);
            var clothingNibcoBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nibco roll" && c.PM_Number == 2);
            var clothingNibcoBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nibco roll" && c.PM_Number == 3);
            var clothingNibcoBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nibco roll" && c.PM_Number == 4);
            var rowNibcoBottom = weeklyPMTable.AddRow();
            rowNibcoBottom.Height = 10;
            var rowNibcoBottom2 = weeklyPMTable.AddRow();
            rowNibcoBottom2.Height = 10;

            var cellNibcoBottom = rowNibcoBottom.Cells[0];
            cellNibcoBottom.Format.Font.Size = 6.5;
            cellNibcoBottom.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom.MergeDown = 1;
            cellNibcoBottom.AddParagraph("NIBCO BOTTOM");

            #region Nibco Bottom PM 1
            // ----- Nibco Bottom Position PM 1 ----- //
            var cellNibcoBottom1Past = rowNibcoBottom.Cells[1];
            cellNibcoBottom1Past.Format.Font.Size = 6.5;
            cellNibcoBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 29 && c.StatusID == 3).Select(c => c.Age).Average();
            nibcoBottom1Past = nibcoBottom1Past != null ? Math.Round((double)nibcoBottom1Past) : 0;
            cellNibcoBottom1Past.AddParagraph($"{nibcoBottom1Past}");

            var cellNibcoBottom1Current = rowNibcoBottom.Cells[2];
            cellNibcoBottom1Current.Format.Font.Size = 6.5;
            cellNibcoBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom1Age = 0;
            if (clothingNibcoBottom1 != null)
            {
                nibcoBottom1Age = clothingNibcoBottom1.Age != null ? Convert.ToInt32(clothingNibcoBottom1.Age) : 0;
            }
            cellNibcoBottom1Current.AddParagraph($"{nibcoBottom1Age}");

            var cellNibcoBottom1Goal = rowNibcoBottom.Cells[3];
            cellNibcoBottom1Goal.Format.Font.Size = 6.5;
            cellNibcoBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottomlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 29);
            cellNibcoBottom1Goal.AddParagraph($"{nibcoBottomlGoal.Goal1}");

            var cellNibcoBottom1Info = rowNibcoBottom2.Cells[1];
            cellNibcoBottom1Info.Format.Font.Size = 6.5;
            cellNibcoBottom1Info.MergeRight = 2;
            cellNibcoBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom1Info.Shading.Color = Colors.LightBlue;
            if (clothingNibcoBottom1 != null)
            {
                cellNibcoBottom1Info.AddParagraph($"{clothingNibcoBottom1.Manufacturer} {clothingNibcoBottom1.Serial_Number}");
            }
            else cellNibcoBottom1Info.AddParagraph("NA");
            #endregion

            #region Nibco Bottom PM 2
            // ----- Nibco Bottom Position PM 2 ----- //
            var cellNibcoBottom2Past = rowNibcoBottom.Cells[4];
            cellNibcoBottom2Past.Format.Font.Size = 6.5;
            cellNibcoBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 29 && c.StatusID == 3).Select(c => c.Age).Average();
            nibcoBottom2Past = nibcoBottom2Past != null ? Math.Round((double)nibcoBottom2Past) : 0;
            cellNibcoBottom2Past.AddParagraph($"{nibcoBottom2Past}");

            var cellNibcoBottom2Current = rowNibcoBottom.Cells[5];
            cellNibcoBottom2Current.Format.Font.Size = 6.5;
            cellNibcoBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom2Age = 0;
            if (clothingNibcoBottom2 != null)
            {
                nibcoBottom2Age = clothingNibcoBottom2.Age != null ? Convert.ToInt32(clothingNibcoBottom2.Age) : 0;
            }
            cellNibcoBottom2Current.AddParagraph($"{nibcoBottom2Age}");

            var cellNibcoBottom2Goal = rowNibcoBottom.Cells[6];
            cellNibcoBottom2Goal.Format.Font.Size = 6.5;
            cellNibcoBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 29);
            cellNibcoBottom2Goal.AddParagraph($"{nibcoBottom2Goal.Goal1}");

            var cellNibcoBottom2Info = rowNibcoBottom2.Cells[4];
            cellNibcoBottom2Info.Format.Font.Size = 6.5;
            cellNibcoBottom2Info.MergeRight = 2;
            cellNibcoBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom2Info.Shading.Color = Colors.LightBlue;
            if (clothingNibcoBottom2 != null)
            {
                cellNibcoBottom2Info.AddParagraph($"{clothingNibcoBottom2.Manufacturer} {clothingNibcoBottom2.Serial_Number}");
            }
            else cellNibcoBottom2Info.AddParagraph("NA");
            #endregion

            #region Nibco Bottom PM 3
            // ----- Nibco Bottom Position PM 3 ----- //
            var cellNibcoBottom3Past = rowNibcoBottom.Cells[7];
            cellNibcoBottom3Past.Format.Font.Size = 6.5;
            cellNibcoBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 29 && c.StatusID == 3).Select(c => c.Age).Average();
            nibcoBottom3Past = nibcoBottom3Past != null ? Math.Round((double)nibcoBottom3Past) : 0;
            cellNibcoBottom3Past.AddParagraph($"{nibcoBottom3Past}");

            var cellNibcoBottom3Current = rowNibcoBottom.Cells[8];
            cellNibcoBottom3Current.Format.Font.Size = 6.5;
            cellNibcoBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom3Age = 0;
            if (clothingNibcoBottom3 != null)
            {
                nibcoBottom3Age = clothingNibcoBottom3.Age != null ? Convert.ToInt32(clothingNibcoBottom3.Age) : 0;
            }
            cellNibcoBottom3Current.AddParagraph($"{nibcoBottom3Age}");

            var cellNibcoBottom3Goal = rowNibcoBottom.Cells[9];
            cellNibcoBottom3Goal.Format.Font.Size = 6.5;
            cellNibcoBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 29);
            cellNibcoBottom3Goal.AddParagraph($"{nibcoBottom3Goal.Goal1}");

            var cellNibcoBottom3Info = rowNibcoBottom2.Cells[7];
            cellNibcoBottom3Info.Format.Font.Size = 6.5;
            cellNibcoBottom3Info.MergeRight = 2;
            cellNibcoBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom3Info.Shading.Color = Colors.LightBlue;
            if (clothingNibcoBottom3 != null)
            {
                cellNibcoBottom3Info.AddParagraph($"{clothingNibcoBottom3.Manufacturer} {clothingNibcoBottom3.Serial_Number}");
            }
            else cellNibcoBottom3Info.AddParagraph("NA");
            #endregion

            #region Nibco Bottom PM 4
            // ----- Nibco Bottom Position PM 4 ----- //
            var cellNibcoBottom4Past = rowNibcoBottom.Cells[10];
            cellNibcoBottom4Past.Format.Font.Size = 6.5;
            cellNibcoBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 29 && c.StatusID == 3).Select(c => c.Age).Average();
            nibcoBottom4Past = nibcoBottom4Past != null ? Math.Round((double)nibcoBottom4Past) : 0;
            cellNibcoBottom4Past.AddParagraph($"{nibcoBottom4Past}");

            var cellNibcoBottom4Current = rowNibcoBottom.Cells[11];
            cellNibcoBottom4Current.Format.Font.Size = 6.5;
            cellNibcoBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom4Age = 0;
            if (clothingNibcoBottom4 != null)
            {
                nibcoBottom4Age = clothingNibcoBottom4.Age != null ? Convert.ToInt32(clothingNibcoBottom4.Age) : 0;
            }
            cellNibcoBottom4Current.AddParagraph($"{nibcoBottom4Age}");

            var cellNibcoBottom4Goal = rowNibcoBottom.Cells[12];
            cellNibcoBottom4Goal.Format.Font.Size = 6.5;
            cellNibcoBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            var nibcoBottom4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 29);
            cellNibcoBottom4Goal.AddParagraph($"{nibcoBottom4Goal.Goal1}");

            var cellNibcoBottom4Info = rowNibcoBottom2.Cells[10];
            cellNibcoBottom4Info.Format.Font.Size = 6.5;
            cellNibcoBottom4Info.MergeRight = 2;
            cellNibcoBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom4Info.Shading.Color = Colors.LightBlue;
            if (clothingNibcoBottom4 != null)
            {
                cellNibcoBottom4Info.AddParagraph($"{clothingNibcoBottom4.Manufacturer} {clothingNibcoBottom4.Serial_Number}");
            }
            else cellNibcoBottom4Info.AddParagraph("NA");
            #endregion

            #endregion // Nibco Bottom Position

            #region Couch Roll Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingCouch1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 1);
            var clothingCouch2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 2);
            var clothingCouch3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 3);
            var clothingCouch4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 4);
            var rowCouch = weeklyPMTable.AddRow();
            rowCouch.Height = 10;
            var rowCouch2 = weeklyPMTable.AddRow();
            rowCouch2.Height = 10;

            var cellCouch = rowCouch.Cells[0];
            cellCouch.Format.Font.Size = 6.5;
            cellCouch.VerticalAlignment = VerticalAlignment.Center;
            cellCouch.MergeDown = 1;
            cellCouch.AddParagraph("COUCH ROLL");

            #region Couch Roll PM 1
            // ----- Couch Position PM 1 ----- //
            var cellCouch1Past = rowCouch.Cells[1];
            cellCouch1Past.Format.Font.Size = 6.5;
            cellCouch1Past.VerticalAlignment = VerticalAlignment.Center;
            var couch1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 30 && c.StatusID == 3).Select(c => c.Age).Average();
            couch1Past = couch1Past != null ? Math.Round((double)couch1Past) : 0;
            cellCouch1Past.AddParagraph($"{couch1Past}");

            var cellCouch1Current = rowCouch.Cells[2];
            cellCouch1Current.Format.Font.Size = 6.5;
            cellCouch1Current.VerticalAlignment = VerticalAlignment.Center;
            var couch1Age = 0;
            if (clothingCouch1 != null)
            {
                couch1Age = clothingCouch1.Age != null ? Convert.ToInt32(clothingCouch1.Age) : 0;
            }
            cellCouch1Current.AddParagraph($"{couch1Age}");

            var cellCouch1Goal = rowCouch.Cells[3];
            cellCouch1Goal.Format.Font.Size = 6.5;
            cellCouch1Goal.VerticalAlignment = VerticalAlignment.Center;
            var couchlGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 30);
            cellCouch1Goal.AddParagraph($"{couchlGoal.Goal1}");

            var cellCouch1Info = rowCouch2.Cells[1];
            cellCouch1Info.Format.Font.Size = 6.5;
            cellCouch1Info.MergeRight = 2;
            cellCouch1Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch1Info.Shading.Color = Colors.LightBlue;
            if (clothingCouch1 != null)
            {
                cellCouch1Info.AddParagraph($"{clothingCouch1.Manufacturer} {clothingCouch1.Serial_Number}");
            }
            else cellCouch1Info.AddParagraph("NA");
            #endregion

            #region Couch Roll PM 2
            // ----- Couch Position PM 2 ----- //
            var cellCouch2Past = rowCouch.Cells[4];
            cellCouch2Past.Format.Font.Size = 6.5;
            cellCouch2Past.VerticalAlignment = VerticalAlignment.Center;
            var couch2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 30 && c.StatusID == 3).Select(c => c.Age).Average();
            couch2Past = couch2Past != null ? Math.Round((double)couch2Past) : 0;
            cellCouch2Past.AddParagraph($"{couch2Past}");

            var cellCouch2Current = rowCouch.Cells[5];
            cellCouch2Current.Format.Font.Size = 6.5;
            cellCouch2Current.VerticalAlignment = VerticalAlignment.Center;
            var couch2Age = 0;
            if (clothingCouch2 != null)
            {
                couch2Age = clothingCouch2.Age != null ? Convert.ToInt32(clothingCouch2.Age) : 0;
            }
            cellCouch2Current.AddParagraph($"{couch2Age}");

            var cellCouch2Goal = rowCouch.Cells[6];
            cellCouch2Goal.Format.Font.Size = 6.5;
            cellCouch2Goal.VerticalAlignment = VerticalAlignment.Center;
            var couch2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 30);
            cellCouch2Goal.AddParagraph($"{couch2Goal.Goal1}");

            var cellCouch2Info = rowCouch2.Cells[4];
            cellCouch2Info.Format.Font.Size = 6.5;
            cellCouch2Info.MergeRight = 2;
            cellCouch2Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch2Info.Shading.Color = Colors.LightBlue;
            if (clothingCouch2 != null)
            {
                cellCouch2Info.AddParagraph($"{clothingCouch2.Manufacturer} {clothingCouch2.Serial_Number}");
            }
            else cellCouch2Info.AddParagraph("NA");
            #endregion

            #region Couch Roll PM 3
            // ----- Couch Position PM 3 ----- //
            var cellCouch3Past = rowCouch.Cells[7];
            cellCouch3Past.Format.Font.Size = 6.5;
            cellCouch3Past.VerticalAlignment = VerticalAlignment.Center;
            var couch3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 30 && c.StatusID == 3).Select(c => c.Age).Average();
            couch3Past = couch3Past != null ? Math.Round((double)couch3Past) : 0;
            cellCouch3Past.AddParagraph($"{couch3Past}");

            var cellCouch3Current = rowCouch.Cells[8];
            cellCouch3Current.Format.Font.Size = 6.5;
            cellCouch3Current.VerticalAlignment = VerticalAlignment.Center;
            var couch3Age = 0;
            if (clothingCouch3 != null)
            {
                couch3Age = clothingCouch3.Age != null ? Convert.ToInt32(clothingCouch3.Age) : 0;
            }
            cellCouch3Current.AddParagraph($"{couch3Age}");

            var cellCouch3Goal = rowCouch.Cells[9];
            cellCouch3Goal.Format.Font.Size = 6.5;
            cellCouch3Goal.VerticalAlignment = VerticalAlignment.Center;
            var couch3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 30);
            cellCouch3Goal.AddParagraph($"{couch3Goal.Goal1}");

            var cellCouch3Info = rowCouch2.Cells[7];
            cellCouch3Info.Format.Font.Size = 6.5;
            cellCouch3Info.MergeRight = 2;
            cellCouch3Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch3Info.Shading.Color = Colors.LightBlue;
            if (clothingCouch3 != null)
            {
                cellCouch3Info.AddParagraph($"{clothingCouch3.Manufacturer} {clothingCouch3.Serial_Number}");
            }
            else cellCouch3Info.AddParagraph("NA");
            #endregion

            #region Couch Roll PM 4
            // ----- Couch Position PM 4 ----- //
            var cellCouch4Past = rowCouch.Cells[10];
            cellCouch4Past.Format.Font.Size = 6.5;
            cellCouch4Past.VerticalAlignment = VerticalAlignment.Center;
            var couch4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 30 && c.StatusID == 3).Select(c => c.Age).Average();
            couch4Past = couch4Past != null ? Math.Round((double)couch4Past) : 0;
            cellCouch4Past.AddParagraph($"{couch4Past}");

            var cellCouch4Current = rowCouch.Cells[11];
            cellCouch4Current.Format.Font.Size = 6.5;
            cellCouch4Current.VerticalAlignment = VerticalAlignment.Center;
            var couch4Age = 0;
            if (clothingCouch4 != null)
            {
                couch4Age = clothingCouch4.Age != null ? Convert.ToInt32(clothingCouch4.Age) : 0;
            }
            cellCouch4Current.AddParagraph($"{couch4Age}");

            var cellCouch4Goal = rowCouch.Cells[12];
            cellCouch4Goal.Format.Font.Size = 6.5;
            cellCouch4Goal.VerticalAlignment = VerticalAlignment.Center;
            var couch4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 30);
            cellCouch4Goal.AddParagraph($"{couch4Goal.Goal1}");

            var cellCouch4Info = rowCouch2.Cells[10];
            cellCouch4Info.Format.Font.Size = 6.5;
            cellCouch4Info.MergeRight = 2;
            cellCouch4Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch4Info.Shading.Color = Colors.LightBlue;
            if (clothingCouch4 != null)
            {
                cellCouch4Info.AddParagraph($"{clothingCouch4.Manufacturer} {clothingCouch4.Serial_Number}");
            }
            else cellCouch4Info.AddParagraph("NA");
            #endregion

            #endregion // Couch Roll Position

            #region L_In Hope Roll Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingLInHope1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 1);
            var clothingLInHope2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 2);
            var clothingLInHope3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 3);
            var clothingLInHope4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 4);
            var rowLInHope = weeklyPMTable.AddRow();
            rowLInHope.Height = 10;
            var rowLInHope2 = weeklyPMTable.AddRow();
            rowLInHope2.Height = 10;

            var cellLInHope = rowLInHope.Cells[0];
            cellLInHope.Format.Font.Size = 6.5;
            cellLInHope.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope.MergeDown = 1;
            cellLInHope.AddParagraph("L-IN HOPE ROLL");

            #region L_In Hope Roll PM 1
            // ----- L_In Hope Roll Position PM 1 ----- //
            var cellLInHope1Past = rowLInHope.Cells[1];
            cellLInHope1Past.Format.Font.Size = 6.5;
            cellLInHope1Past.VerticalAlignment = VerticalAlignment.Center;
            var lInHope1Past = db.Clothings.Where(c => c.PM_Number == 1 && c.PositionID == 31 && c.StatusID == 3).Select(c => c.Age).Average();
            lInHope1Past = lInHope1Past != null ? Math.Round((double)lInHope1Past) : 0;
            cellLInHope1Past.AddParagraph($"{lInHope1Past}");

            var cellLInHope1Current = rowLInHope.Cells[2];
            cellLInHope1Current.Format.Font.Size = 6.5;
            cellLInHope1Current.VerticalAlignment = VerticalAlignment.Center;
            var lInHope1Age = 0;
            if (clothingLInHope1 != null)
            {
                lInHope1Age = clothingLInHope1.Age != null ? Convert.ToInt32(clothingLInHope1.Age) : 0;
            }
            cellLInHope1Current.AddParagraph($"{lInHope1Age}");

            var cellLInHope1Goal = rowLInHope.Cells[3];
            cellLInHope1Goal.Format.Font.Size = 6.5;
            cellLInHope1Goal.VerticalAlignment = VerticalAlignment.Center;
            var lInHopelGoal = db.Goals.SingleOrDefault(g => g.PM_ID == 1 && g.PositionID == 31);
            cellLInHope1Goal.AddParagraph($"{lInHopelGoal.Goal1}");

            var cellLInHope1Info = rowLInHope2.Cells[1];
            cellLInHope1Info.Format.Font.Size = 6.5;
            cellLInHope1Info.MergeRight = 2;
            cellLInHope1Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope1Info.Shading.Color = Colors.LightBlue;
            if (clothingLInHope1 != null)
            {
                cellLInHope1Info.AddParagraph($"{clothingLInHope1.Manufacturer} {clothingLInHope1.Serial_Number}");
            }
            else cellLInHope1Info.AddParagraph("NA");
            #endregion

            #region L_In Hope Roll PM 2
            // ----- L_In Hope Roll Position PM 2 ----- //
            var cellLInHope2Past = rowLInHope.Cells[4];
            cellLInHope2Past.Format.Font.Size = 6.5;
            cellLInHope2Past.VerticalAlignment = VerticalAlignment.Center;
            var lInHope2Past = db.Clothings.Where(c => c.PM_Number == 2 && c.PositionID == 31 && c.StatusID == 3).Select(c => c.Age).Average();
            lInHope2Past = lInHope2Past != null ? Math.Round((double)lInHope2Past) : 0;
            cellLInHope2Past.AddParagraph($"{lInHope2Past}");

            var cellLInHope2Current = rowLInHope.Cells[5];
            cellLInHope2Current.Format.Font.Size = 6.5;
            cellLInHope2Current.VerticalAlignment = VerticalAlignment.Center;
            var lInHope2Age = 0;
            if (clothingLInHope2 != null)
            {
                lInHope2Age = clothingLInHope2.Age != null ? Convert.ToInt32(clothingLInHope2.Age) : 0;
            }
            cellLInHope2Current.AddParagraph($"{lInHope2Age}");

            var cellLInHope2Goal = rowLInHope.Cells[6];
            cellLInHope2Goal.Format.Font.Size = 6.5;
            cellLInHope2Goal.VerticalAlignment = VerticalAlignment.Center;
            var lInHope2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 31);
            cellLInHope2Goal.AddParagraph($"{lInHope2Goal.Goal1}");

            var cellLInHope2Info = rowLInHope2.Cells[4];
            cellLInHope2Info.Format.Font.Size = 6.5;
            cellLInHope2Info.MergeRight = 2;
            cellLInHope2Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope2Info.Shading.Color = Colors.LightBlue;
            if (clothingLInHope2 != null)
            {
                cellLInHope2Info.AddParagraph($"{clothingLInHope2.Manufacturer} {clothingLInHope2.Serial_Number}");
            }
            else cellLInHope2Info.AddParagraph("NA");
            #endregion

            #region L_In Hope Roll PM 3
            // ----- L_In Hope Roll Position PM 3 ----- //
            var cellLInHope3Past = rowLInHope.Cells[7];
            cellLInHope3Past.Format.Font.Size = 6.5;
            cellLInHope3Past.VerticalAlignment = VerticalAlignment.Center;
            var lInHope3Past = db.Clothings.Where(c => c.PM_Number == 3 && c.PositionID == 31 && c.StatusID == 3).Select(c => c.Age).Average();
            lInHope3Past = lInHope3Past != null ? Math.Round((double)lInHope3Past) : 0;
            cellLInHope3Past.AddParagraph($"{lInHope3Past}");

            var cellLInHope3Current = rowLInHope.Cells[8];
            cellLInHope3Current.Format.Font.Size = 6.5;
            cellLInHope3Current.VerticalAlignment = VerticalAlignment.Center;
            var lInHope3Age = 0;
            if (clothingLInHope3 != null)
            {
                lInHope3Age = clothingLInHope3.Age != null ? Convert.ToInt32(clothingLInHope3.Age) : 0;
            }
            cellLInHope3Current.AddParagraph($"{lInHope3Age}");

            var cellLInHope3Goal = rowLInHope.Cells[9];
            cellLInHope3Goal.Format.Font.Size = 6.5;
            cellLInHope3Goal.VerticalAlignment = VerticalAlignment.Center;
            var lInHope3Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 3 && g.PositionID == 31);
            cellLInHope3Goal.AddParagraph($"{lInHope3Goal.Goal1}");

            var cellLInHope3Info = rowLInHope2.Cells[7];
            cellLInHope3Info.Format.Font.Size = 6.5;
            cellLInHope3Info.MergeRight = 2;
            cellLInHope3Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope3Info.Shading.Color = Colors.LightBlue;
            if (clothingLInHope3 != null)
            {
                cellLInHope3Info.AddParagraph($"{clothingLInHope3.Manufacturer} {clothingLInHope3.Serial_Number}");
            }
            else cellLInHope3Info.AddParagraph("NA");
            #endregion

            #region L_In Hope Roll PM 4
            // ----- L_In Hope Roll Position PM 4 ----- //
            var cellLInHope4Past = rowLInHope.Cells[10];
            cellLInHope4Past.Format.Font.Size = 6.5;
            cellLInHope4Past.VerticalAlignment = VerticalAlignment.Center;
            var lInHope4Past = db.Clothings.Where(c => c.PM_Number == 4 && c.PositionID == 31 && c.StatusID == 3).Select(c => c.Age).Average();
            lInHope4Past = lInHope4Past != null ? Math.Round((double)lInHope4Past) : 0;
            cellLInHope4Past.AddParagraph($"{lInHope4Past}");

            var cellLInHope4Current = rowLInHope.Cells[11];
            cellLInHope4Current.Format.Font.Size = 6.5;
            cellLInHope4Current.VerticalAlignment = VerticalAlignment.Center;
            var lInHope4Age = 0;
            if (clothingLInHope4 != null)
            {
                lInHope4Age = clothingLInHope4.Age != null ? Convert.ToInt32(clothingLInHope4.Age) : 0;
            }
            cellLInHope4Current.AddParagraph($"{lInHope4Age}");

            var cellLInHope4Goal = rowLInHope.Cells[12];
            cellLInHope4Goal.Format.Font.Size = 6.5;
            cellLInHope4Goal.VerticalAlignment = VerticalAlignment.Center;
            var lInHope4Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 4 && g.PositionID == 31);
            cellLInHope4Goal.AddParagraph($"{lInHope4Goal.Goal1}");

            var cellLInHope4Info = rowLInHope2.Cells[10];
            cellLInHope4Info.Format.Font.Size = 6.5;
            cellLInHope4Info.MergeRight = 2;
            cellLInHope4Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope4Info.Shading.Color = Colors.LightBlue;
            if (clothingLInHope4 != null)
            {
                cellLInHope4Info.AddParagraph($"{clothingLInHope4.Manufacturer} {clothingLInHope4.Serial_Number}");
            }
            else cellLInHope4Info.AddParagraph("NA");
            #endregion

            #endregion // L_In Hope Roll Position

            #region L_Out Hope Roll Position

            #endregion // L_Out Hope Roll Position

            #region Bottom Press Wringer Position

            #endregion // Bottom Press Wringer Position

            #region Top Press Wringer Position

            #endregion // Top Press Wringer Position

            #region Couch Paper Roll Position

            #endregion // Couch Paper Roll Position

            return weeklyPMTable;
        }
    }
}