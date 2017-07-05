using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//Dynamo 
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;

using EtabsConnection;
using ETABS2016;

namespace DynamoETABS.Assembly
{
    public class Read
    {
        /// DYNAMO NODES 
        
        /// <summary>
        /// Read ETABS project from a filepath 
        /// </summary> 
        /// <param name="FilePath">File path of the project to read</param>
        /// <param name="read">Set Boolean to True to open and read the project</param>
        /// <returns>Structural Model</returns>
        [MultiReturn("StructuralModel", "units")]
        public static Dictionary<string, object> ETABSModel(String Filepath, bool read)
        {
            if(read)
            {
                //StructuralModel Model = new StructuralModel Model();
                //Model.StructuralElements = new List<Element>();
                cSapModel myEtabsModel = null;
                string units = string.Empty;
                //Open & Initialize ETABS File 
                Initialize.OpenEtabsModel(Filepath, ref myEtabsModel, ref units);

                //Populate the model's elements 
                //StructuralModelFromETABFIle(ref myEtabsModel, ref Model, units);

                //Return outputs
                return new Dictionary<string, object>
                {
                    {"StructuralModel", 0},
                    {"units", units}
                };
            }
            else
            {
                throw new Exception("Set boolean True to Read!");
            }
        }


    }
}
