/// Develop by Edbert Widjaja - HERA ENGINEERING 
/// 5.07.17 
/// 


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//ETABS
using ETABS2016;
//Interop.COM services for ETABS 
using System.Runtime.InteropServices;

//DYNAMO
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using System.Diagnostics;
using System.Collections;

namespace EtabsConnection
{
    public class Initialize
    {
        public static void InitializeSapModel(ref cOAPI myETABSObject, ref cSapModel myEtabsModel, string units)
        {

            long ret = 0;

            //Create API helper object
            ETABS2016.cHelper myHelper;
            myHelper = new ETABS2016.Helper();

            //Create ETABS2016 Object 
            myETABSObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");

            //get enum 
            eUnits Units = (eUnits)Enum.Parse(typeof(eUnits), units);

            //Start ETABS Application
            ret = myETABSObject.ApplicationStart();

            // Create EtabsModel Object
            //Get a reference to cSapModel to Access all API Classes and functions 
            myEtabsModel = default(ETABS2016.cSapModel);
            myEtabsModel = myETABSObject.SapModel;

            //Initialize model 
            ret = myEtabsModel.InitializeNewModel();

            //Create a new blank model 
            ret = myEtabsModel.File.NewBlank();

        }

        public static String GetModelFilename (ref cSapModel myEtabsModel, ref string units)
        {
            return myEtabsModel.GetModelFilename();
        }

        public static void OpenEtabsModel(String filepath, ref cSapModel myEtabsModel, ref string units)
        {
            long ret = 0;
            //Create ETABS Object 

            //Create API helper object
            ETABS2016.cHelper myHelper;
            myHelper = new ETABS2016.Helper();

            //Create ETABS2016 Object
            ETABS2016.cOAPI myEtabsObject;
            myEtabsObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");

            //Start Application 
            myEtabsObject.ApplicationStart();

            //Create Etabs Model 
            myEtabsModel = default(ETABS2016.cSapModel);
            myEtabsModel = myEtabsObject.SapModel;
            ret = myEtabsModel.InitializeNewModel();
            ret = myEtabsModel.File.OpenFile(filepath);
            units = myEtabsModel.GetPresentUnits().ToString(); 
        }

        public static void GrabOpenETABS(ref cSapModel myEtabsModel, ref string ModelUnits, string DynInputUnits = "kN_m_C")
        {
            Process[] SapInstances = Process.GetProcessesByName("ETABS 2016");

            if (SapInstances.LongLength >= 1)
            {
                ETABS2016.cOAPI myEtabsObject = null;
                myEtabsObject = (ETABS2016.cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject");
                
            }

            long ret = 0;

      

        }











    }
}
