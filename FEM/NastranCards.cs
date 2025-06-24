using Microsoft.Office.Interop.Excel;

namespace StressUtilities.FEM
{
    class NastranCards
    {
        private string StartCell = "B3";

        public void WriteNastranCards(string LoadComponent)
        {
            //string LoadComponent = "FORCE";
            NastranLoadCards(LoadComponent);
        }

        private void NastranLoadCards(string LoadComponent)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet CardSheet;
            string NasCards="";
            if (WorksheetExists(LoadComponent, wb))
            {
                CardSheet = wb.Worksheets[LoadComponent];
            }
            else
            {
                wb.Worksheets.Add();
                wb.ActiveSheet.Name = LoadComponent;
                CardSheet = wb.Worksheets[LoadComponent];
            }


            switch (LoadComponent)
            {
                case "FORCE":
                    //FORCE SID G CID F N1 N2 N3
                    NasCards = "FORCE;SID;G;CID;F;N1;N2;N3";
                    break;
                case "FORCE1":
                    //FORCE1 SID G F G1 G2
                    NasCards = "FORCE1;SID;G;F;G1;G2";
                    break;
                case "FORCE2":
                    //FORCE2 SID G F G1 G2 G3 G4
                    NasCards = "FORCE2;SID;G;F;G1;G2;G3;G4";
                    break;
                case "MOMENT":
                    //MOMENT SID G CID M N1 N2 N3
                    NasCards = "MOMENT;SID;G;CID;M;N1;N2;N3";
                    break;
                case "MOMENT1":
                    //MOMENT1 SID G M G1 G2
                    NasCards = "MOMENT1;SID;G;M;G1;G2";
                    break;
                case "MOMENT2":
                    //MOMENT2 SID G M G1 G2 G3 G4
                    NasCards = "MOMENT2;SID;G;M;G1;G2;G3;G4";
                    break;
                case "PLOAD":
                    //PLOAD SID P G1 G2 G3 G4
                    NasCards = "PLOAD;SID;P;G1;G2;G3;G4";
                    break;
                case "PLOAD1":
                    //PLOAD1 SID EID TYPE SCALE X1 P1 X2 P2
                    NasCards = "PLOAD1;SID;EID;TYPE;SCALE;X1;P1;X2;P2";
                    break;
                case "PLOAD2":
                    //PLOAD2 SID P EID1 EID2 EID3 EID4 EID5 EID6 EID7 EID8 -etc.-
                    NasCards = "PLOAD2;SID;P;EID1;EID2;EID3;EID4;EID5;EID6;EID7;EID8;-etc.-";
                    break;
                case "PLOADB3":
                    //PLOADB3 SID EID CID N1 N2 N3 TYPE SCALE P(A) P(B) P(C)
                    NasCards = "PLOADB3;SID;EID;CID;N1;N2;N3;TYPE;SCALE;P(A);P(B);P(C)";
                    break;
                case "PLOAD4":
                    //PLOAD4 SID EID P1 P2 P3 P4 G1 (G3 or G4) CID N1 N2 N3 SORL LDIR
                    NasCards = "PLOAD4;SID;EID;P1;P2;P3;P4;G1;(G3 or G4);CID;N1;N2;N3;SORL;LDIR";
                    break;
                case "PLOADX1":
                    //PLOADX1 SID EID PA PB GA GB THETA
                    NasCards = "PLOADX1;SID;EID;PA;PB;GA;GB;THETA";
                    break;
                case "PRESAX":
                    NasCards = "PRESAX;SID;P;RID1;RID2;PHI1;PHI2";
                    break;
                case "SLOAD":  //Also with SOL153 and SOL400
                    //SLOAD SID S1 F1 S2 F2 S3 F3
                    NasCards = "SLOAD;SID;S1;F1;S2;F2;S3;F3";
                    break;
                case "RFORCE":
                    //RFORCE SID G CID A R1 R2 R3 METHOD RACC MB IDRF
                    NasCards = "RFORCE;SID;G;CID;A;R1;R2;R3;METHOD;RACC;MB;IDRF";
                    break;
                case "GRAV":
                    //GRAV SID CID A N1 N2 N3 MB
                    NasCards = "GRAV;SID;CID;A;N1;N2;N3;MB";
                    break;
                case "ACCEL":
                    //ACCEL SID CID N1 N2 N3 DIR LOC1 VAL1 LOC2 VAL2 (Continues in Groups of 2)
                    NasCards = "ACCEL;SID;CID;N1;N2;N3;DIR;LOC1;VAL1;LOC2;VAL2;(Continues in Groups of 2)";
                    break;
                case "ACCEL1":
                    //ACCEL1 SID CID A N1 N2 N3 GRIDID1 GRIDID2 -etc.-
                    NasCards = "ACCEL1;SID;CID;A;N1;N2;N3;GRIDID1;GRIDID2;-etc.-";
                    break;
                case "SPCD":  //SOL600 only
                    //SPCD SID G1 C1 D1 G2 C2 D2
                    NasCards = "SPCD;SID;G1;C1;D1;G2;C2;D2";
                    break;
                case "SPCR": //SOL600 only
                    //SPCR SID G1 C1 D1 G2 C2 D2
                    NasCards = "SPCR;SID;G1;C1;D1;G2;C2;D2";
                    break;
                case "QBDY1": //SOL153 and SOL400 only  -- Heat Flux
                    //QBDY1 SID Q0 EID1 EID2 EID3 EID4 EID5 EID6 EID7 EID8 -etc.- 
                    NasCards = "QBDY1;SID;Q0;EID1;EID2;EID3;EID4;EID5;EID6;EID7;EID8;-etc.-";
                    break;
                case "QBDY2":
                    //QBDY2 SID EID Q01 Q02 Q03 Q04 Q05 Q06 Q07 Q08
                    NasCards = "QBDY2;SID;EID;Q01;Q02;Q03;Q04;Q05;Q06;Q07;Q08";
                    break;
                case "QBDY3":
                    //QBDY3 SID Q0 CNTRLND EID1 EID2 EID3 EID4 EID5 EID6 etc.
                    NasCards = "QBDY3;SID;Q0;CNTRLND;EID1;EID2;EID3;EID4;EID5;EID6;etc.";
                    break;
                case "QVECT":
                    //QVECT SID Q0 TSOUR CE (E1 or TID1) (E2 or TID2) (E3 or TID3) CNTRLND EID1 EID2 -etc.-
                    NasCards = "QVECT;SID;Q0;TSOUR;CE;(E1 or TID1);(E2 or TID2);(E3 or TID3);CNTRLND;EID1;EID2;-etc.-";
                    break;
                case "QVOL":
                    //QVOL SID QVOL CNTRLND EID1 EID2 EID3 EID4 EID5 EID6 etc.
                    NasCards = "QVOL;SID;QVOL;CNTRLND;EID1;EID2;EID3;EID4;EID5;EID6;etc.";
                    break;
                case "QHBDY":
                    //QHBDY SID FLAG Q0 AF G1 G2 G3 G4 G5 G6 G7 G8
                    NasCards = "QHBDY;SID;FLAG;Q0;AF;G1;G2;G3;G4;G5;G6;G7;G8";
                    break;
                case "DAREA":  //if these entries have been converted
                    //DAREA SID P1 C1 A1 P2 C2 A2
                    NasCards = "DAREA;SID;P1;C1;A1;P2;C2;A2";
                    break;
                case "TEMP":
                    //TEMP SID G1 T1 G2 T2 G3 T3
                    NasCards = "TEMP;SID;G1;T1;G2;T2;G3;T3";
                    break;
                case "TEMPD":
                    NasCards = "TEMPD;SID1;T1;SID2;T2;SID3;T3;SID4;T4";
                    break;
                case "TEMPAX":
                    NasCards = "TEMPAX;SID1;RID1;PHI1;T1;SID2;RID2;PHI2;T2";
                    break;
                case "TEMPBC":
                    NasCards = "TEMPBC;SID;TYPE;TEMP1;GID1;TEMP2;GID2;TEMP3;GID3";
                    break;
                case "SPC":
                    NasCards = "SPC;SID;G1;C1;D1;G2;C2;D2";
                    break;
                case "SPC1":
                    NasCards = "SPC1;SID;C;G1;G2;G3;G4;G5;G6;G7;G8;G9;-etc.-";
                    break;
                case "SPCAX":
                    NasCards = "SPCAX;SID;RID;HID;C;D";
                    break;
                case "DEFORM":
                    NasCards = "DEFORM;SID;EID1;D1;EID2;D2;EID3;D3";
                    break;
                case "RLOAD1":
                    NasCards = "RLOAD1;SID;EXCITEID;DELAYI/DELAYR;DPHASEI/DPHASER;TC/RC;TD/RD;TYPE";
                    break;
                case "RLOAD2":
                    NasCards = "RLOAD2;SID;EXCITEID;DELAYI/DELAYR;DPHASEI/DPHASER;TB/RB;TP/RP;TYPE";
                    break;
                case "DLOAD":
                    NasCards = "DLOAD;SID;S;S1;L1;S2;L2;S3;L3;S4;L4;-etc.-;";
                    break;
                case "TLOAD1":
                    NasCards = "TLOAD1;SID;EXCITEID;DELAYI/DELAYR;TYPE;TID/F;US0;VS0";
                    break;
                case "TLOAD2":
                    NasCards = "TLOAD2;SID;EXCITEID;DELAYI/DELAYR;TYPE;T1;T2;F;P;C;B;US0;VS0";
                    break;
                case "SUBCASE":
                    NasCards = "SUBCASE;TITLE;SUBTITLE;LABEL;LOAD;SPC;DISP;SPCFORCE;FORCE;STRESS;TEMPERATURE(LOAD);OLOAD;DEFORM;MPC;DISPLACEMENT;SET;MODES";
                    break;
                case "LOAD":
                    NasCards = "LOAD;SID;S;S1;L1;S2;L2;S3;L3;S4;L4;-etc.-;";
                    break;
                case "LOADT":
                    NasCards = "LOADT;SID;L1;T1;L2;T2;L3;T3;L4;T4;L5;T5;-etc.-;";
                    break;
            }
            PopulateNastranCards(NasCards, CardSheet);

            Range Selection = CardSheet.Range[CardSheet.Range[StartCell], CardSheet.Range[StartCell].End[XlDirection.xlToRight].Offset[10, 0]];
            Selection.Borders.LineStyle = XlLineStyle.xlContinuous;
        }


        private void PopulateNastranCards(string NasCards, Worksheet CardSheet)
        {
            string[] HeadingArray;
            long RowNdx, ColNdx;
            Range Selection;
            HeadingArray = NasCards.Split(';');
            Selection=CardSheet.Range[StartCell];
            RowNdx = Selection.Row;
            ColNdx = Selection.Column;

            Selection.Offset[-1, 0].Value = "NOTE: THE TABLE MUST BEGIN AT THE CELL B3";
            Selection.Offset[-1, 0].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); ;
            for (int i=0;i<HeadingArray.Length;i++)
            {
                CardSheet.Cells[RowNdx, ColNdx + i].value = HeadingArray[i];
            }
        }

        private bool WorksheetExists(string WorksheetName, Workbook wb)
        {
            Application xlApp = Globals.ThisAddIn.Application;

            bool sheetExists = false;
            //Worksheet Sht;
            foreach (Worksheet Sht in wb.Worksheets)
            {
                //if (xlApp.Application.Proper(Sht.Name) == xlApp.Application.Proper(WorksheetName))
                if (Sht.Name == WorksheetName)
                {
                    return true;
                }
            }
            return sheetExists;

        }

    }
}
