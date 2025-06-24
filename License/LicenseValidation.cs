using Portable.Licensing;
using Portable.Licensing.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace StressUtilities
{
    class LicenseValidation
    {

        private Portable.Licensing.License ULicense
        { get; set; }
        private string PublicKey
        {
            get
            {
                return @"MIIBKjCB4wYHKoZIzj0CATCB1wIBATAsBgcqhkjOPQEBAiEA/////wAAAAEAAAAAAAAAAAAAAAD///////////////8wWwQg/////wAAAAEAAAAAAAAAAAAAAAD///////////////wEIFrGNdiqOpPns+u9VXaYhrxlHQawzFOw9jvOPD4n0mBLAxUAxJ02CIbnBJNqZnjhE50mt4GffpAEIQNrF9Hy4SxCR/i85uVjpEDydwN9gS3rM6D0oTlF2JjClgIhAP////8AAAAA//////////+85vqtpxeehPO5ysL8YyVRAgEBA0IABIhMwNqkMO1TEN1QGa/iG+q7ZcWvlo5FHR7eNqiYdSWwNBVzM89YcfcPVNm40D7fb2lShPA9gBLL9NvFboYzDQM=";
            }
        }

        private string ValidateLicense(Portable.Licensing.License license)
        {
            //IValidationFailure  validationFailure;
            string ReturnValue = "License is Valid";

            IEnumerable<IValidationFailure> validationFailures = license.Validate().ExpirationDate().When(LicenseException).And().Signature(PublicKey).AssertValidLicense();

            if (validationFailures.Any())//(validationFailures != null)
            {
                ReturnValue = "";
                try
                {
                    foreach (IValidationFailure validationFailure in validationFailures)
                    {
                        ReturnValue += $"{validationFailure.HowToResolve} : \n validationFailure.Message \n";
                    }
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }

            return ReturnValue;

        }

        private bool LicenseException(Portable.Licensing.License license)
        {
            if (license.Type == LicenseType.Trial)
                return true;

            return false;
        }

        public string LicenseValidityCheck()
        {
            Portable.Licensing.License license;
            license = GetLicenseStream();

            return ValidateLicense(license);
        }

        private Portable.Licensing.License GetLicenseStream()
        {
            Stream myStream = null;
            string LicensePath;
            string LicenseFileName;
            bool GetLicFileFlag = true;

            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //Location Is where the assembly Is run from 
            string assemblyLocation = assemblyInfo.Location;

            //CodeBase Is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

            LicensePath = ClickOnceLocation;
            LicenseFileName = LicensePath + @"\License_StressUtilities.lic";

            if (File.Exists(LicenseFileName))
            {
                myStream = new FileStream(LicenseFileName, FileMode.Open);
                ULicense = Portable.Licensing.License.Load(myStream);

                if (ValidateLicense(ULicense) == "License is Valid")
                {
                    GetLicFileFlag = false;
                }

                if (myStream != null)
                {
                    myStream.Close();
                }

            }
            return ULicense;
        }

        public void ImportLicense()
        {
            Stream myStream = null;
            string LicensePath;
            bool GetLicFileFlag = true;
            string LicenseFileName;
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //Location Is where the assembly Is run from 
            string assemblyLocation = assemblyInfo.Location;

            //CodeBase Is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

            LicensePath = ClickOnceLocation;
            LicenseFileName = LicensePath + "\\" + "License_StressUtilities.lic";


            if (GetLicFileFlag == true)
            {

                using (OpenFileDialog OfDLicense = new OpenFileDialog())
                {
                    OfDLicense.Filter = @"License (*.lic)|*.lic";
                    OfDLicense.FilterIndex = 1;
                    OfDLicense.RestoreDirectory = true;


                    if (OfDLicense.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            myStream = OfDLicense.OpenFile();

                            if (myStream != null)
                            {
                                ULicense = Portable.Licensing.License.Load(myStream);

                                // Code to save the filestream to the system.
                                File.Copy(OfDLicense.FileName, LicenseFileName, true);
                                MessageBox.Show("License Successfully Imported. Please restart Excel to validate the license");
                            }
                        }
                        catch (Exception Ex)
                        {
                            MessageBox.Show($"Cannot read file from disk. Original error: {Ex.Message}");
                        }
                        finally
                        {
                            // Check this again, since we need to make sure we didn't throw an exception on open.
                            if (myStream != null)
                            {
                                myStream.Close();
                            }
                        }
                    }
                }
            }
        }
    }
}
