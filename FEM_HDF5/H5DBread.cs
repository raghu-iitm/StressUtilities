using HDF.PInvoke;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
//using hid_t = System.Int64;
using herr_t = System.Int32;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using StressUtilities;
using System.Globalization;
using System.Threading;
using System.Linq;
//using Application = System.Windows.Forms.Application;

/** 
Copyright (c) 2020-2030 Raghavendra Prasad Laxman
Licensed under the GPL-3.0 license. See LICENSE file for details.
*/


#if HDF5_VER1_10
using hid_t = System.Int64;
#else
using hid_t = System.Int32;
#endif
namespace Nastranh5
{
    public class H5DBread
    {
        private CultureInfo cSystemCulture = Thread.CurrentThread.CurrentCulture;
        private bool _disposed;
        private long HDF5DataLimit = (long)StressUtilities.Properties.Settings.Default.MaxHDFRows;

        public H5DBread()
        {
        }

        public void Dispose()
        {
            // Dispose of unmanaged resources.
            Dispose(true);
            // Suppress finalization.
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                // TODO: dispose managed state (managed objects).
            }

            // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
            // TODO: set large fields to null.

            _disposed = true;
        }


        public void LaunchHDF5Form()
        {
            IEnumerable<Nash5> FrmCollection = Application.OpenForms.OfType<Nash5>();
            if (FrmCollection.Any())
                FrmCollection.First().Focus();
            else
            {
                Nash5 HDF5Form = new Nash5();
                HDF5Form.Show();
            }
        }

        private string GetAsciiString(string unicodeString)
        {

            // Create two different encodings.
            Encoding ascii = Encoding.ASCII;
            Encoding unicode = Encoding.Unicode;

            // Convert the string into a byte array.
            byte[] unicodeBytes = unicode.GetBytes(unicodeString);

            // Perform the conversion from one encoding to the other.
            byte[] asciiBytes = Encoding.Convert(unicode, ascii, unicodeBytes);

            // Convert the new byte[] into a char[] and then into a string.
            char[] asciiChars = new char[ascii.GetCharCount(asciiBytes, 0, asciiBytes.Length)];
            ascii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0);
            string asciiString = new string(asciiChars);
            return asciiString;
        }

        public void ExtractH5File(ref List<string> h5FileList, string grpName, ref string nasDataset, ref string datasetparent,
      ref List<long> entitylist, ref List<long> SubCaseList, bool[] Requests, ref string StartRange, ref string LocRequest, ref bool Success)
        {
            hid_t dataspace;
            int rank;
            herr_t status_n;
            ulong[] NULL = null, dims;
            long cparms;
            hid_t memspace;
            string grpIndex = null;
            Dictionary<string, Dictionary<string, string>> IndexTable = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, Dictionary<string, string>> Domains = new Dictionary<string, Dictionary<string, string>>();
            string grpDomain = @"/NASTRAN/RESULT";
            string scase = null;
            string firstkey = null;
            bool Resultdataset = false;

            List<string> HeaderList = new List<string>();
            //long RowNdx, ColNdx, StartRow;
            H5General gencls = new H5General();

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            string DatasetType = gencls.GetDatasetType(nasDataset); // replace the code to work with xml data.

            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = wb.ActiveSheet;
            Excel.Range Selection = null;

            string ofilePath = wb.Path;

            if (!string.IsNullOrEmpty(ws.Range[StartRange].Text))
            {
                DialogResult Result = MessageBox.Show(@"The target cell is not empty. Are you sure to continue?...", "Warning!", MessageBoxButtons.YesNo);
                if (Result == DialogResult.No)
                {
                    Marshal.ReleaseComObject(ws);
                    Marshal.ReleaseComObject(wb);
                    return;
                }
                else
                {
                    ws.Range[StartRange].ClearContents();
                    Selection = ws.Range[ws.Range[StartRange].Offset[1,0], ws.Range[StartRange].Offset[1,0].End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                    Selection.ClearContents();
                    Selection.ClearFormats();
                    Marshal.ReleaseComObject(Selection);
                }
            }

            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            xlApp.ScreenUpdating = false;
            ws.Range[StartRange].Value2 = grpName+"/"+nasDataset;
            
            long RowNdx = ws.Range[StartRange].Row+1;
            long ColNdx = ws.Range[StartRange].Column;
            long StartRow = RowNdx;

            hid_t faplist_id = (hid_t)H5P.create(H5P.FILE_ACCESS);

            herr_t status = H5P.set_fapl_stdio(faplist_id);

            bool condwriteHead = true;
            for (int fc = 0; fc < h5FileList.Count; fc++)
            {
                hid_t fileId = (hid_t)H5F.open(GetAsciiString(h5FileList[fc]), H5F.ACC_RDONLY, H5P.DEFAULT); //faplist_id
                if (fileId <= 0)
                {
                    MessageBox.Show(@"Cannot open hdf5 file ");

                    xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    xlApp.ScreenUpdating = true;
                    H5F.close(fileId);
                    H5.close();
                    Marshal.ReleaseComObject(Selection);
                    Marshal.ReleaseComObject(ws);
                    Marshal.ReleaseComObject(wb);

                    return;
                }
                hid_t groupID = H5G.open(fileId, grpName, H5P.DEFAULT);
                hid_t dsetId = H5D.open(groupID, nasDataset, H5P.DEFAULT);
                hid_t typeID = H5D.get_type(dsetId);

                dataspace = H5D.get_space(dsetId);
                H5T.order_t order = H5T.get_order(typeID);

                rank = H5S.get_simple_extent_ndims(dataspace);
                dims = new ulong[rank];

                status_n = H5S.get_simple_extent_dims(dataspace, dims, NULL);

                cparms = H5D.get_create_plist(dsetId); // /* Get properties handle first. */

                herr_t nmembers = H5T.get_nmembers(typeID);

                H5T.class_t typcls = H5T.get_class(typeID);

                IntPtr sizestr = H5T.get_size(typeID);

                hid_t mem_s1_t = H5T.get_native_type(typeID, H5T.direction_t.DEFAULT);

                hid_t plist_id = H5D.get_access_plist(dsetId);

                ulong offstdb = H5D.get_offset(dsetId);

                // Define the memory space to read dataset.

                memspace = H5S.create_simple(rank, dims, NULL);
                // Dim resultdata As New typcls
                hid_t npoints = H5S.get_select_npoints(dataspace);

                hid_t sizexx = H5S.get_select_npoints(memspace);

                hid_t grpDomID = H5G.open(fileId, grpDomain, H5P.DEFAULT);
                hid_t domaindsetId = H5D.open(grpDomID, "DOMAINS", H5P.DEFAULT);
                Domains = IndexData(domaindsetId); //, "DOMAINS"
                H5D.close(domaindsetId);
                H5G.close(grpDomID);

                foreach (string key in Domains.Keys)
                {
                    firstkey = key;
                    break;
                }
                if (grpName.Contains(grpDomain) && grpName != grpDomain && grpName.StartsWith("/NASTRAN/RESULT"))
                {
                    Resultdataset = true;
                }
                if (Domains[firstkey].ContainsKey("SUBCASE") && Resultdataset == true)
                {
                    scase = "SUBCASE";
                }

                // Write File Headers********************
                if (condwriteHead)
                {
                    //  Header to be combined based on dataset(if 2D or 3D)

                    HeaderList = WriteHeaders(nmembers, typeID, ref Requests, ref DatasetType, ref scase);

                    int offset = 0;

                    Selection = ws.Range[ws.Cells[RowNdx, ColNdx], ws.Cells[RowNdx, ColNdx].Offset[0, HeaderList.Count - 1]];
                    Selection.ClearFormats();
                    Selection.Value = HeaderList.ToArray();
                    Selection.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    offset = HeaderList.Count;
                    RowNdx++;
                    condwriteHead = false;

                } // End of File Headers

                //Insert a loop for indexing here
                if (grpName.StartsWith("/NASTRAN/RESULT"))
                {
                    grpIndex = string.Concat("/INDEX", grpName);

                }
                hid_t grpNdxID = H5G.open(fileId, grpIndex, H5P.DEFAULT);
                hid_t grpdsetId = H5D.open(grpNdxID, nasDataset, H5P.DEFAULT);


                if (grpdsetId >= 0)
                {
                    try
                    {
                        IndexTable = IndexData(grpdsetId); //nasDataset
                        H5D.close(grpdsetId);
                        H5G.close(grpNdxID);


                        StartRow = RowNdx - 1;
                        string Pos, len;
                        ulong Poshs, lenhs;
                        foreach (string domainID in IndexTable.Keys)
                        {
                            Pos = IndexTable[domainID]["POSITION"];
                            len = IndexTable[domainID]["LENGTH"];
                            Poshs = Convert.ToUInt64(Pos);
                            lenhs = Convert.ToUInt64(len);

                            scase = Domains[domainID]["SUBCASE"];
                            if (grpName.StartsWith("/NASTRAN/RESULT"))
                            {
                                if (SubCaseList.Count != 0 && scase != null && !SubCaseList.Contains(long.Parse(scase)))
                                {
                                    continue;
                                }
                            }

                                if (RowNdx - StartRow < HDF5DataLimit)
                                ReadHDF5Memory(ref Poshs, ref lenhs, dataspace, nmembers, sizestr, dsetId, typeID, order, dims,
                            ref entitylist, ref SubCaseList, ref HeaderList, ref DatasetType, ref datasetparent, ref Requests,
                            ref scase, ref LocRequest, ref grpName, ref ofilePath, ref Success, ref ws, ref RowNdx, ref ColNdx, ref StartRow, ref xlApp);

                        }
                    }
                    catch (Exception ex)
                    {
                        Success = false;
                        H5D.close(grpdsetId);
                        H5G.close(grpNdxID);
                        H5S.close(memspace);
                        H5P.close(plist_id);
                        H5D.close(typeID);
                        H5D.close(dsetId);
                        H5T.close(mem_s1_t);
                        H5T.close(nmembers);
                        H5S.close(dataspace);
                        H5G.close(groupID);
                        H5F.close(fileId);
                        H5P.close(faplist_id);
                        H5.close();
                        H5.garbage_collect();
                        Marshal.ReleaseComObject(Selection);
                        Marshal.ReleaseComObject(ws);
                        Marshal.ReleaseComObject(wb);

                        xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                        xlApp.ScreenUpdating = true;
                        MessageBox.Show(ex.ToString());
                        return;
                    }
                }
                else
                {
                    H5D.close(grpdsetId);
                    H5G.close(grpNdxID);
                    ulong Pos = 0;
                    ulong len = dims[0];
                    scase = null;
                    ReadHDF5Memory(ref Pos, ref len, dataspace, nmembers, sizestr, dsetId, typeID, order, dims,
                    ref entitylist, ref SubCaseList, ref HeaderList, ref DatasetType, ref datasetparent, ref Requests,
                    ref scase, ref LocRequest, ref grpName, ref ofilePath, ref Success, ref ws, ref RowNdx, ref ColNdx, ref StartRow, ref xlApp);
                }
                //Success = true;
                Domains = null;
                IndexTable = null;
                H5S.close(memspace);
                H5D.close(typeID);
                H5P.close(plist_id);
                H5D.close(dsetId);
                H5T.close(mem_s1_t);
                H5T.close(nmembers);
                H5S.close(dataspace);  //H5D.Close or H5S.Close?
                H5G.close(groupID);
                H5F.close(fileId);
                H5P.close(faplist_id);
                H5.close();
                H5.garbage_collect();

                Marshal.ReleaseComObject(Selection);


            }

            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            xlApp.ScreenUpdating = true;

            Thread.CurrentThread.CurrentCulture = new CultureInfo(cSystemCulture.Name);
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void ReadHDF5Memory(ref ulong Pos, ref ulong len, hid_t dataspace, hid_t nmembers, IntPtr sizestr,
            hid_t dsetId, hid_t typeID, H5T.order_t order, ulong[] dims, ref List<hid_t> entitylist, ref List<long>  SubCaseList,
            ref List<string> HeaderList, ref string DatasetType, ref string datasetparent, ref bool[] Requests, ref string scase, ref string LocRequest,
            ref string grpName, ref string ofilePath, ref bool Success, ref Excel.Worksheet ws, ref long RowNdx, ref long ColNdx, ref long StartRow, ref Excel.Application xlApp)
        {

            int statusread; //herr_t
            ulong[] NULL = null;
            List<object> ResultData = new List<object>();
            List<object[]> ResultDataSummary = new List<object[]>();
            object[,] ResultDataSummaryArray = new object[1, 1];
            List<object>[] FinalResult = new List<object>[0];

            Excel.Range Selection = null;

            hid_t memspace;
            H5T.class_t memcls;
            IntPtr ptr = IntPtr.Zero, ofstPointer = IntPtr.Zero, bufferptr = IntPtr.Zero;
            GCHandle handle;
            List<string> dataset = new List<string>();
            List<object> outList = new List<object>();
            //List<List<object>> strContainer = new List<List<object>>();
            string[] strArray = { };
            object[,] ResultMatrix = null;
            hid_t Printentity = 0;
            int i = 0, listcount = 0, iStore = 0, finalvalue;
            string strvalue = null, ColKey;

            hid_t mem_type_ID, size;

            byte[] buffer;
            ulong[] offset = { Pos }, stride = { 1 }, count = { len }, block = { 1 }, dimsm = { len };

            herr_t status = H5S.select_hyperslab(dataspace, H5S.seloper_t.SET, offset, stride, count, NULL);

            bool boolWriteResult;
            int ofstPtr, Nodef;

            //Define memory dataspace
            offset[0] = 0;
            stride[0] = 1;
            count[0] = len;
            block[0] = 1;

            memspace = H5S.create_simple(1, dimsm, NULL);
            hid_t status_m = H5S.select_hyperslab(memspace, H5S.seloper_t.SET, offset, stride, count, NULL);

            hid_t sizestrint = sizestr.ToInt64();

            byte[] resultdata = new byte[sizestrint * (int)len];
            handle = GCHandle.Alloc(resultdata, GCHandleType.Pinned);
            //ptr = handle.AddrOfPinnedObject();

            statusread = H5D.read(dsetId, typeID, memspace, dataspace, H5P.DEFAULT, handle.AddrOfPinnedObject()); // TypeID

            handle.Free();

            hid_t mem_type = H5T.copy(typeID);


            bool arrayResult = false;


            bool writelog = false;
            StringBuilder MissingIds = new StringBuilder("");
            double RequestRatio = (double)entitylist.Count / (double)len;

            while (i < (int)len)  //|| i<=HDF5DataLimit
            {
                //xlApp.StatusBar = $"Progress: {(double)i/(double)len*100:0.00}%";
                //get element position position
                if (entitylist.Count != 0 && RequestRatio <= 0.05)  //5% cutoff to decide between linear and binary search.
                {
                    iStore = i;
                    i = SearchIndex(ref entitylist, mem_type, ref resultdata, i, (int)len, ref listcount, (int)sizestrint);
                    if (i == -1)
                    {
                        if (listcount < entitylist.Count)
                        {
                            Success = false;
                            StrConcat(MissingIds, entitylist[listcount - 1].ToString());
                            i = iStore;
                            writelog = true;
                            continue;
                        }
                        else if (listcount == entitylist.Count)
                        {
                            StrConcat(MissingIds, entitylist[listcount - 1].ToString());
                            writelog = true;
                            break;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                outList = new List<object>();
                ResultMatrix = null;
                boolWriteResult = true;
                //ofstPtr;
                Nodef = 1;

                for (uint k = 0; k < nmembers; k++)
                {
                    strvalue = null;
                    memcls = H5T.get_member_class(typeID, k);
                    mem_type_ID = H5T.get_member_type(mem_type, k);

                    size = H5T.get_size(mem_type_ID).ToInt64();

                    ofstPointer = H5T.get_member_offset(mem_type, k);
                    ofstPtr = ofstPointer.ToInt32();

                    buffer = new byte[size];
                    handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);

                    Marshal.Copy(resultdata, i * (int)sizestrint + ofstPtr, handle.AddrOfPinnedObject(), (int)size);
                    handle.Free();

                    strvalue = Get_Values(memcls, mem_type_ID, order, buffer, ofstPtr, size, dims);

                    H5T.close(mem_type_ID);

                    /*  switch (SortOption)
                      {
                          case "SORTED":
                              break;

                          case "UNSORTED":
                              break;

                      }*/

                    ColKey = Marshal.PtrToStringAnsi(H5T.get_member_name(typeID, k));

                    if (ColKey == "NODEF")
                    {
                        Nodef = Convert.ToInt32(strvalue);
                    }

                    if (ofstPtr == 0 && boolWriteResult == true)
                    {
                        if (Int64.TryParse(strvalue, out Int64 _))
                        {
                            if (!entitylist.Contains(Convert.ToInt64(strvalue)) && entitylist.Count != 0)
                            {
                                boolWriteResult = false;
                                break;
                            }
                            else
                            {
                                Printentity = Convert.ToInt64(strvalue);
                            }
                        }
                    }

                    if (memcls == H5T.class_t.ARRAY)
                    {
                        if (grpName.StartsWith("/NASTRAN/RESULT"))
                        {
                            strArray = strvalue.Split(',');
                            if (LocRequest == "Centroid")
                            {
                                arrayResult = false;
                                strvalue = strArray[0];
                            }
                            else
                            {
                                arrayResult = true;
                            }
                        }
                        else
                        {
                            strvalue = strvalue.Replace(',', ' '); // Replace the heading with the suffix and split the string. The succeeding 0s in the Grid field.
                        }
                    }

                    if (arrayResult == true)
                    {
                        if (ResultMatrix == null)   // Imporve the code since the code is very slow.
                        {
                            ResultMatrix = new object[strArray.Length, nmembers];
                            for (int m = 0; m < strArray.Length; ++m)
                            {
                                for (int n = 0; n < k; n++)
                                {
                                    ResultMatrix[m, n] = outList[n];
                                }
                                ResultMatrix[m, k] = strArray[m];
                            }
                        }
                        else
                        {
                            if (memcls == H5T.class_t.ARRAY)
                                for (int m = 0; m < strArray.Length; ++m)
                                {
                                    ResultMatrix[m, k] = strArray[m];
                                }
                            else
                                for (int m = 0; m < strArray.Length; ++m)
                                {
                                    ResultMatrix[m, k] = strvalue;
                                }
                        }

                    }
                    else
                    {
                        outList.Add(strvalue);
                    }
                }


                if (grpName.StartsWith("/NASTRAN/RESULT"))
                {
                    //if(SubCaseList.Count!=0 && scase !=null && !SubCaseList.Contains(long.Parse(scase)))
                    //{
                    //    ++i;
                    //    continue;
                    //}

                    if (!arrayResult)
                    {
                        if (scase != null && datasetparent != "RESULT")
                        {
                            outList.Insert(1, scase);
                        }
                        CalculateDerivedLoads(ref HeaderList, ref outList, ref DatasetType, ref datasetparent, ref Requests);
                    }
                    else
                    {
                        //FinalResult = DeriveEquivalent(ref HeaderList, ref ResultMatrix, ref DatasetType, ref datasetparent, ref Requests);
                    }


                }
                //strContainer = null;

                if (boolWriteResult)
                {
                    if (ResultDataSummary.Count < 10000)
                        if (!arrayResult)
                            ResultDataSummary.Add(Array.ConvertAll(outList.ToArray(), s => double.TryParse(s.ToString(), out double xresult) ? xresult : s));
                        else
                        {

                            finalvalue = ResultMatrix.GetLength(0);

                            for (int j = 1; j < finalvalue; j++)
                            {
                                outList = new List<object>();

                                for (int k = 0; k < ResultMatrix.GetLength(1); k++)
                                {
                                    outList.Add(ResultMatrix[j, k]);
                                }
                                //outList.InsertRange(outList.Count - 1, FinalResult[j].Cast<object>().ToArray());
                                if (scase != null && datasetparent != "RESULT")
                                {
                                    outList.Insert(1, scase);
                                }
                                CalculateDerivedLoads(ref HeaderList, ref outList, ref DatasetType, ref datasetparent, ref Requests);

                                ResultDataSummary.Add(Array.ConvertAll(outList.ToArray(), s => double.TryParse(s.ToString(), out double xresult) ? xresult : s));
                            }
                            ResultMatrix = null;
                            arrayResult = false;
                        }
                    else
                    {
                        if (!arrayResult)
                            ResultDataSummary.Add(Array.ConvertAll(outList.ToArray(), s => double.TryParse(s.ToString(), out double xresult) ? xresult : s));
                        else
                        {

                            finalvalue = ResultMatrix.GetLength(0);
                            for (int j = 1; j < finalvalue; j++)
                            {
                                outList = new List<object>();
                                for (int k = 0; k < ResultMatrix.GetLength(1); k++)
                                {
                                    outList.Add(ResultMatrix[j, k]);
                                }
                                //outList.InsertRange(outList.Count - 1, FinalResult[j].Cast<object>().ToArray());
                                if (scase != null && datasetparent != "RESULT")
                                {
                                    outList.Insert(1, scase);
                                }
                                
                                CalculateDerivedLoads(ref HeaderList, ref outList, ref DatasetType, ref datasetparent, ref Requests);
                                ResultDataSummary.Add(Array.ConvertAll(outList.ToArray(), s => double.TryParse(s.ToString(), out double xresult) ? xresult : s));
                            }
                            ResultMatrix = null;
                            arrayResult = false;
                        }
                        //if (ResultDataSummaryArray.GetLength(0) != ResultDataSummary.Count)
                        ResultDataSummaryArray = new object[ResultDataSummary.Count, outList.Count];
                        for (int row = 0; row < ResultDataSummary.Count; row++)
                        {
                            for (int col = 0; col < ResultDataSummary[row].Length; col++)
                            {
                                ResultDataSummaryArray[row, col] = ResultDataSummary[row][col];
                            }
                        }
                        Selection = ws.Range[ws.Cells[RowNdx, ColNdx], ws.Cells[RowNdx, ColNdx].Offset[ResultDataSummary.Count - 1, outList.Count - 1]];
                        Selection.ClearFormats();
                        Selection.Value2 = ResultDataSummaryArray;
                        Selection.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        Marshal.ReleaseComObject(Selection);
                        RowNdx += ResultDataSummary.Count;
                        ResultDataSummary.Clear();
                        ResultDataSummaryArray = null;
                    }

                    if (RowNdx - StartRow >= HDF5DataLimit)
                    {
                        H5S.close(status_m);
                        outList.Clear();
                        ResultDataSummaryArray = null;
                        ptr = IntPtr.Zero;
                        Marshal.ReleaseComObject(Selection);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        MessageBox.Show($"Limitation of {HDF5DataLimit} lines has been reached. The data extraction stopped");
                        return;
                    }
                }

                if (listcount >= entitylist.Count && entitylist.Count != 0)
                {
                    break;
                }
                if (entitylist.Count == 0 || (entitylist.Count != 0 && RequestRatio > 0.05))
                {
                    ++i;
                }

            }     //End of while loop

            H5T.close(mem_type);

            if (ResultDataSummary.Count != 0)
            {
                ResultDataSummaryArray = new object[ResultDataSummary.Count, outList.Count];
                for (int row = 0; row < ResultDataSummary.Count; row++)
                {
                    for (int col = 0; col < ResultDataSummary[row].Length; col++)
                    {
                        ResultDataSummaryArray[row, col] = ResultDataSummary[row][col];
                    }
                }
                Selection = ws.Range[ws.Cells[RowNdx, ColNdx], ws.Cells[RowNdx, ColNdx].Offset[ResultDataSummary.Count - 1, outList.Count - 1]];
                Selection.ClearFormats();
                Selection.Value2 = ResultDataSummaryArray;
                Selection.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                Selection.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                RowNdx += ResultDataSummary.Count;
                ResultDataSummary.Clear();
                outList.Clear();
                ptr = IntPtr.Zero;
            }

            if (writelog)
            {
                StreamWriter tlog = new StreamWriter(ofilePath + "_" + "log" + ".log");
                tlog.WriteLine("The following entities are not found in the dataset");
                tlog.WriteLine(MissingIds.ToString());
                tlog.Close();
            }

            outList.Clear();
            H5S.close(status_m);
            H5.garbage_collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private int SearchIndex(ref List<hid_t> entitylist, hid_t mem_type, ref byte[] resultdata, int i, int len, ref int listcount, int sizestrint)
        {
            GCHandle handle;

            hid_t base_ID = H5T.get_member_type(mem_type, 0);
            hid_t basesize = H5T.get_size(base_ID).ToInt64();

            byte[] Posnbuffer = new byte[basesize];
            handle = GCHandle.Alloc(Posnbuffer, GCHandleType.Pinned);
            int Start = i;
            int End = len - 1;
            int mid;
            hid_t target;

            H5T.close(base_ID);

            while (Start <= End)
            {
                mid = (Start + End) / 2;

                Marshal.Copy(resultdata, mid * (int)sizestrint, handle.AddrOfPinnedObject(), (int)basesize);

                target = BitConverter.ToInt64(Posnbuffer, 0);

                if (target == entitylist[listcount])
                {
                    i = mid;
                    listcount++;
                    handle.Free();
                    return i;
                }
                else
                {
                    if (target < entitylist[listcount])
                    {
                        Start = mid + 1;
                    }

                    if (target > entitylist[listcount])
                    {
                        End = mid - 1;
                    }
                }
            }
            handle.Free();
            listcount++;
            return -1;

        }

        /* private void CentroidValues(ref string[] strArray, ref object[,] ResultMatrix, ref List<object> outList, ref string DatasetType,
             ref string datasetparent, string ColKey, ref int Nodef, ref string LocRequest)
         {
             //string temp = "test";
             switch (DatasetType)
             {
                 case "SS2D":
                 case "CPLXSS2D":
                 case "SS3D":
                 case "CPLXSS3D":
                     if (datasetparent == "STRESS" || datasetparent == "STRAIN")
                     {
                         if (ColKey == "GRID")
                         {
                             //StrConcat(outString, Convert.ToString(strArray[0]));
                             outList.Add(Convert.ToString(strArray[0]));
                         }
                         else
                         {
                             double Sum = 0.0;
                             double Average = 0.0;
                             for (int m = 1; m < strArray.Length; ++m)
                             {
                                 Sum = Sum + Convert.ToDouble(strArray[m]);
                             }
                             Average = Sum / Nodef;

                             outList.Add(Convert.ToString(Average));
                         }
                     }

                     break;
                 default:
                     if (ResultMatrix.GetLength(0) != strArray.Length)
                     {
                         //strContainer = new List<List<object>>();
                         List<object> ResultData = new List<object>();
                         ResultData = outList;
                         ResultData.AddRange(strArray);
                         //strContainer.Add(ResultData);

                     }
                     else
                     {
                         List<object> ResultData = new List<object>();
                         ResultData = outList;
                         ResultData.AddRange(strArray);
                     }
                     break;

             }

             //return temp;

         }*/

        private Dictionary<string, Dictionary<string, string>> IndexData(hid_t dsetId) //, string nasDataset
        {
            hid_t dataspace;
            int rank;
            herr_t status_n;
            int statusread; //herr_t
            ulong[] NULL = null;
            long cparms;
            ulong[] dims;
            hid_t memspace;
            H5T.class_t memcls;
            IntPtr ptr;
            GCHandle handle;
            List<string> dataset = new List<string>();
            string[] strContainer = { };
            string[] strArray = { };
            Dictionary<string, Dictionary<string, string>> IndexDict = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, string> offsetDict = new Dictionary<string, string>();

            hid_t typeID = H5D.get_type(dsetId);

            dataspace = H5D.get_space(dsetId);
            H5T.order_t order = H5T.get_order(typeID);

            rank = H5S.get_simple_extent_ndims(dataspace);
            dims = new ulong[rank];

            status_n = H5S.get_simple_extent_dims(dataspace, dims, NULL);   // H5T.str_t.NULLTERM

            cparms = H5D.get_create_plist(dsetId); // /* Get properties handle first. */

            herr_t nmembers = H5T.get_nmembers(typeID);

            H5T.class_t typcls = H5T.get_class(typeID);

            var sizestr = H5T.get_size(typeID);


            hid_t mem_s1_t = H5T.get_native_type(typeID, H5T.direction_t.DEFAULT);  // H5T.DIR_DEFAULT
                                                                                    // Dim plist_id As hssize_t = H5D.get_create_plist(dsetId)
            hid_t plist_id = H5D.get_access_plist(dsetId);

            ulong offstdb = H5D.get_offset(dsetId);

            // Define the memory space to read dataset.

            memspace = H5S.create_simple(rank, dims, NULL);

            hid_t npoints = H5S.get_select_npoints(dataspace);

            hid_t sizexx = H5S.get_select_npoints(memspace);

            ulong[] offset = { 0 };
            ulong[] stride = { 1 };
            ulong[] count = { dims[0] };
            ulong[] dimsm = { dims[0] };
            ulong[] block = { 1 };

            H5S.close(memspace);

            hid_t status = H5S.select_hyperslab(dataspace, H5S.seloper_t.SET, offset, stride, count, NULL);

            //Define memory dataspace
            memspace = H5S.create_simple(1, dimsm, NULL);
            hid_t status_m = H5S.select_hyperslab(memspace, H5S.seloper_t.SET, offset, stride, count, NULL);

            var sizestrint = sizestr.ToInt64();

            byte[] resultdata = new byte[sizestrint * (int)dims[0]];
            handle = GCHandle.Alloc(resultdata, GCHandleType.Pinned);
            ptr = handle.AddrOfPinnedObject();

            statusread = H5D.read(dsetId, typeID, memspace, dataspace, H5P.DEFAULT, ptr); // TypeID
            handle.Free();

            hid_t mem_type = H5T.copy(typeID);

            string strvalue = null, str;
            string[] Headerlist = new string[nmembers];
            string[] IndexValues = new string[nmembers];

            for (uint k = 0; k <= nmembers - 1; k++)
            {
                str = Marshal.PtrToStringAnsi(H5T.get_member_name(typeID, k));
                Headerlist[k] = str;
            }
            IndexDict = new Dictionary<string, Dictionary<string, string>>();
            int ofstPtr;
            IntPtr ofstPointer;
            hid_t mem_type_ID, size;
            byte[] buffer;
            for (int i = 0; i < (int)dims[0]; ++i)
            {


                for (uint k = 0; k <= nmembers - 1; k++)
                {
                    strvalue = null;
                    memcls = H5T.get_member_class(typeID, k);
                    mem_type_ID = H5T.get_member_type(mem_type, k);

                    ofstPointer = H5T.get_member_offset(mem_type, k);
                    ofstPtr = ofstPointer.ToInt32();

                    size = H5T.get_size(mem_type_ID).ToInt64();

                    buffer = new byte[size];
                    handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);

                    //byte[] buffer = dataBytes.Skip<byte>(i * (int)sizestrint).Take<byte>((int)sizestrint).ToArray<byte>();

                    Marshal.Copy(resultdata, i * (int)sizestrint + ofstPtr, handle.AddrOfPinnedObject(), (int)size);
                    handle.Free();


                    strvalue = Get_Values(memcls, mem_type_ID, order, buffer, ofstPtr, size, dims);

                    H5T.close(mem_type_ID);


                    IndexValues[k] = strvalue;

                    //ofstPtr = ofstPtr + (int)size;

                }

                //Dictionary to collect dataset
                offsetDict = new Dictionary<string, string>();

                for (int k = (int)nmembers - 1; k >= 0; k--)
                {

                    if (k == 0)
                    {
                        IndexDict.Add(IndexValues[k], offsetDict);
                    }
                    else
                    {
                        offsetDict.Add(Headerlist[k], IndexValues[k]);
                    }

                }
                offsetDict = null;
            }

            H5T.close(mem_type);
            H5T.close(mem_s1_t);
            H5D.close(typeID);
            H5D.close(dsetId);
            H5S.close(dataspace);
            H5S.close(memspace);
            H5.garbage_collect();
            return IndexDict;
        }

        /*private List<object>[] DeriveEquivalent(ref List<string> HeaderList, ref object[,] ResultMatrix, ref string DatasetType, ref string datasetparent, ref bool[] Requests)
        {
            Tensors CalcTensor = new Tensors();
            Dictionary<string, string> strD = new Dictionary<string, string>();
            object[] vonMises = null;
            object[] Principal = null;
            List<object>[] FinalResult = new List<object>[ResultMatrix.GetLength(0)];

            for (int i = 0; i < ResultMatrix.GetLength(0); i++)
            {
                FinalResult[i] = new List<object>();
                strD = new Dictionary<string, string>();
                for (int j = 0; j < ResultMatrix.GetLength(1); j++)
                {
                    strD.Add(HeaderList[j], ResultMatrix[i, j].ToString());
                }
                switch (DatasetType)
                {
                    case "SS2D":
                        if (Requests[0])
                        {
                            if (datasetparent == "STRESS")
                                Principal = CalcTensor.Principal2D(strD);
                            else
                                Principal = CalcTensor.PrincipalStrain2D(strD);
                        }
                        if (Requests[1])
                        {
                            vonMises = CalcTensor.vonMises2D(strD);
                        }
                        break;
                    case "CPLXSS2D":
                        if (Requests[0])
                            Principal = CalcTensor.vonMises2DCPLX(strD);
                        break;
                    case "SS3D":
                        if (Requests[0])
                        {
                            if (datasetparent == "STRESS")
                                Principal = CalcTensor.Principal3D(strD);
                            else
                                Principal = CalcTensor.PrincipalStrain3D(strD);
                        }

                        if (Requests[1])
                            vonMises = CalcTensor.vonMises3D(strD);
                        break;
                    case "CPLXSS3D":
                        if (Requests[1])
                            vonMises = CalcTensor.vonMises3DCPLX(strD);
                        break;
                    case "NONE":
                        break;
                }

                if (Requests[0])
                {
                    (FinalResult[i]).AddRange(Principal);
                }

                if (Requests[1])
                {
                    FinalResult[i].AddRange(vonMises);
                }

                if (scase != null && datasetparent != "RESULT")
                {
                    FinalResult[i].Add(scase);
                }
            }

            return FinalResult;


        }*/

        private void CalculateDerivedLoads(ref List<string> HeaderList, ref List<object> outList, ref string DatasetType, ref string datasetparent, ref bool[] Requests)
        {
            Tensors CalcTensor = new Tensors();
            Dictionary<string, string> strD = new Dictionary<string, string>();
            object[] vonMises = null;
            object[] Principal = null;

            for (int i = 0; i < outList.Count; i++)
            {
                strD.Add(HeaderList[i], outList[i].ToString());
            }
            switch (DatasetType)
            {
                case "SS2D":
                    if (Requests[0])
                    {
                        if (datasetparent == "STRESS")
                            Principal = CalcTensor.Principal2D(strD);
                        else
                            Principal = CalcTensor.PrincipalStrain2D(strD);
                    }
                    if (Requests[1])
                    {
                        vonMises = CalcTensor.vonMises2D(strD);
                    }
                    break;
                case "CPLXSS2D":
                    if (Requests[0])
                        Principal = CalcTensor.vonMises2DCPLX(strD);
                    break;
                case "SS3D":
                    if (Requests[0])
                    {
                        if (datasetparent == "STRESS")
                            Principal = CalcTensor.Principal3D(strD);
                        else
                            Principal = CalcTensor.PrincipalStrain3D(strD);
                    }
                    if (Requests[1])
                        vonMises = CalcTensor.vonMises3D(strD);
                    break;
                case "CPLXSS3D":
                    if (Requests[1])
                        vonMises = CalcTensor.vonMises3DCPLX(strD);
                    break;
                case "NONE":
                    break;
            }

            if (Requests[0])
            {
                outList.InsertRange(outList.Count - 1, Principal);
            }
            if (Requests[1])
            {
                outList.InsertRange(outList.Count - 1, vonMises);
            }

            /*if (scase != null && datasetparent != "RESULT")
            {
                outList.Add(scase);
            }*/
        }



        private hid_t Get_memsize(hid_t mem_type_ID)
        {
            IntPtr sizeData = H5T.get_size(mem_type_ID);

            hid_t size = sizeData.ToInt64();  //* (int)dims[0]

            return size;
        }



        private string Get_Values(H5T.class_t memcls, hid_t mem_type_ID, H5T.order_t order, byte[] buffer, int ofstPtr, hid_t size, ulong[] dims)
        {
            //Type dataType;

            switch (memcls)
            {
                case H5T.class_t.INTEGER:
                    {
                        return BitConverter.ToInt64(buffer, 0).ToString();
                        // break;
                    }

                case H5T.class_t.FLOAT:
                    {
                        if (size == 4)
                            return BitConverter.ToSingle(buffer, 0).ToString();
                        else if (size == 8)
                            return BitConverter.ToDouble(buffer, 0).ToString();
                        else return null;
                    }

                case H5T.class_t.STRING:
                    {
                        var cSet = H5T.get_cset(mem_type_ID);
                        if (cSet == H5T.cset_t.ASCII)
                        {
                            return System.Text.Encoding.ASCII.GetString(buffer).Trim(); //.TrimEnd('\0')
                        }
                        else
                        {
                            return System.Text.Encoding.UTF8.GetString(buffer).Trim();
                        }

                    }

                case H5T.class_t.BITFIELD:
                    {
                        return BitConverter.ToBoolean(buffer, 0).ToString();
                    }

                case H5T.class_t.OPAQUE:
                    {
                        return null;

                    }

                case H5T.class_t.COMPOUND:
                    {
                        return null;
                    }

                case H5T.class_t.REFERENCE:
                    {
                        return null;
                    }

                case H5T.class_t.ENUM:
                    {
                        //H5T.get_member_value(mem_type_ID, 0, buffer);
                        return null;
                    }

                case H5T.class_t.VLEN:
                    {
                        return null;
                    }

                case H5T.class_t.ARRAY:
                    {
                        herr_t aDim = H5T.get_array_ndims(mem_type_ID);
                        IntPtr sz2 = H5T.get_size(mem_type_ID); //IntPtr
                        herr_t aDim1 = H5T.get_array_ndims(mem_type_ID);
                        hid_t ArrayBaseID = H5T.get_super(mem_type_ID);
                        IntPtr ElemSize = H5T.get_size(ArrayBaseID);
                        hid_t Esize = ElemSize.ToInt64();
                        H5T.class_t ArrCls = H5T.get_class(ArrayBaseID);
                        StringBuilder str = new StringBuilder();

                        for (int i = 0; i < size / Esize; ++i)
                        {
                            byte[] resultdata = new byte[Esize];
                            Array.Copy(buffer, i * Esize, resultdata, 0, Esize);
                            string ResVal = null;
                            switch (ArrCls)
                            {
                                case H5T.class_t.INTEGER:
                                    ResVal = BitConverter.ToInt64(resultdata, 0).ToString();
                                    break;
                                case H5T.class_t.FLOAT:
                                    if (Esize == 4)
                                        ResVal = BitConverter.ToSingle(resultdata, 0).ToString();
                                    else if (Esize == 8)
                                        ResVal = BitConverter.ToDouble(resultdata, 0).ToString();
                                    else ResVal = null;
                                    break;
                                case H5T.class_t.STRING:
                                    ResVal = System.Text.Encoding.ASCII.GetString(resultdata).TrimEnd('\0');
                                    break;
                                default:
                                    ResVal = "Error";
                                    break;
                            }
                            StrConcat(str, ResVal);
                        }
                        H5T.close(ArrayBaseID);
                        return str.ToString();
                    }

                default:
                    {
                        return null;
                    }
            }//End of Switch
        }// End of Function

        private List<string> WriteHeaders(hid_t nmembers, hid_t typeID, ref bool[] Requests, ref string DatasetType, ref string scase)
        {
            List<string> HeaderList = new List<string>();

            for (uint k = 0; k < nmembers; k++)
            {
                string str = Marshal.PtrToStringAnsi(H5T.get_member_name(typeID, k));
                HeaderList.Add(str);
            }

            if (Requests[0])
            {
                switch (DatasetType)
                {
                    case "SS2D": //case "CPLXSS2D":
                        HeaderList.InsertRange(HeaderList.Count - 1, new string[] { "MAX_PRINCIPAL1", "MIN_PRINCIPAL1", "ANGLE1", "MAX_PRINCIPAL2", "MIN_PRINCIPAL2", "ANGLE2" });
                        /*HeaderList.Add("MAX_PRINCIPAL1");
                        HeaderList.Add("MIN_PRINCIPAL1");
                        HeaderList.Add("ANGLE1");
                        HeaderList.Add("MAX_PRINCIPAL2");
                        HeaderList.Add("MIN_PRINCIPAL2");
                        HeaderList.Add("ANGLE2");*/
                        break;
                    case "SS3D": //case "CPLXSS3D":
                        HeaderList.InsertRange(HeaderList.Count - 1, new string[] { "MAX_PRINCIPAL", "MID_PRINCIPAL", "MIN_PRINCIPAL" });
                        /*HeaderList.Add("MAX_PRINCIPAL");
                        HeaderList.Add("MID_PRINCIPAL");
                        HeaderList.Add("MIN_PRINCIPAL");*/
                        break;
                    default:
                        break;
                }
            }

            if (Requests[1])
            {
                HeaderList.Insert(HeaderList.Count - 1, "VON MISES");
            }

            if (scase != null)
            {
                HeaderList.Insert(1, "SUBCASE");
            }
            return HeaderList;
        }

        private void StrConcat(StringBuilder str, string ResVal)
        {
            string listseparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            if (str.Length == 0)
                str.Append(ResVal);
            else
            {
                //str.Append(listseparator);
                str.Append(listseparator + ResVal);
            }

        }


    }
}
