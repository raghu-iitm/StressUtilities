using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using HDF.PInvoke;

#if HDF5_VER1_10
using hid_t = System.Int64;
#else
using hid_t = System.Int32;
#endif

namespace Nastranh5
{
    public class TreeView
    {
        private readonly char delim = '/';
        public Dictionary<string, NodeEntry> GetGroups(List<string> datasetNames)  //, TreeNode nodeToAddTo
        {
            NodeEntryCollection cItems = new NodeEntryCollection();
            string[] TreeList;
            for (int i = 0; i < datasetNames.Count - 1; i++)
            {
                TreeList = datasetNames[i].Split(delim);
                cItems.AddNodeEntry(TreeList, 0);
                //cItems.AddEntry(datasetNames[i], 0);
            }
            return cItems;
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

        public List<string> GetGroupList(string filename)
        {
            hid_t fileId;           /* Handle */

            // Open file.
            try
            {
                fileId = H5F.open(GetAsciiString(filename), H5F.ACC_RDONLY);  //, H5P.DEFAULT
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Error: {ex.ToString()}");
                return null;
            }

                List<string> datasetNames = new List<string>();
                //List<string> groupNames = new List<string>();
                //List<string> NamedDatatype = new List<string>();
                hid_t rootId = H5G.open(fileId, "/");

                int status=H5O.visit(fileId, H5.index_t.NAME, H5.iter_order_t.NATIVE, new H5O.iterate_t(
                    delegate (long objectId, IntPtr namePtr, ref H5O.info_t info, IntPtr op_data)
                    {
                        string objectName = Marshal.PtrToStringAnsi(namePtr);
                        H5O.info_t gInfo = new H5O.info_t();
                        int statusInfo=H5O.get_info_by_name(objectId, objectName, ref gInfo);

                        if (gInfo.type == H5O.type_t.DATASET)
                        {
                            datasetNames.Add(objectName);
                        }
                    /*else if (gInfo.type == H5O.type_t.GROUP)
                    {
                        groupNames.Add(objectName);
                    }
                    else if (gInfo.type == H5O.type_t.NAMED_DATATYPE)
                    {
                        NamedDatatype.Add(objectName);
                    }*/
                        return 0;
                    }), new IntPtr());

                H5G.close(rootId);
                H5F.close(fileId);
                H5.close();
                H5.garbage_collect();
                return datasetNames;
            /*}
            catch(Exception ex)
            {
                MessageBox.Show($"Error: {ex.ToString()}");
                return null;
            }*/

        }


    }
}
