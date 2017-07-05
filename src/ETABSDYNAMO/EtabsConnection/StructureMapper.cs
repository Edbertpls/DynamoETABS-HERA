using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ETABS2016;

//DYNAMO
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime; 

namespace ETABSConnection
{
    [SupressImportIntoVM]
    public class StructureMapper
    {
        // Draw Frame Object return ID 
        public static void CreateUpdateFrm(ref cSapModel Model, double iX, double iY, double iZ, double jX, double jY, double jZ, ref string Id, bool update, ref string error)
        {
            if(!update)
            {
                //1. Create Frame 
                string dummy = string.Empty;
                long ret = Model.FrameObj.AddByCoord(iX, iY, iZ, jX, jY, jZ, ref dummy);
                Id = dummy;
                if (ret == 1) error = string.Format("Error creating frame{0}", dummy);
            }
            else
            {
                //update location if coordiantes have been chagned
            }
        }

        public static void CreateorUpdateArea(ref cSapModel Model, Mesh m, ref string Id, bool update, double SF, ref string error)
        {
            int counter = 0; 
            if(!update) //Create a new one? 
            {
                List<string> ProfilePts = new List<string>();

                long ret = 0; 
                foreach (var v in m.VertexPositions)
                {
                    string dummy = null;
                    ret = Model.PointObj.AddCartesian(v.X * SF, v.Y * SF, v.Z * SF, ref dummy);

                    ProfilePts.Add(dummy); 
                }

                string[] names = ProfilePts.ToArray();
                string dummyarea = string.Empty;
                ret = Model.AreaObj.AddByPoint(ProfilePts.Count(), ref names, ref dummyarea);
                if (ret == 1) counter++;
                Id = dummyarea;
            }
            else
            {
                //Update area code 
            }

            if (counter > 0) error = string.Format("Error creating Mesh{0}",Id);
        }
        public static void CreateorUpdateArea(ref cSapModel Model, Surface s, ref string Id, bool update, double SF, ref string Error)
        {
            Curve[] PerimeterCrvs = s.PerimeterCurves();
            List<Point> SurfPoints = new List<Point>();
            long ret = 0;
            int counter = 0;
            foreach (var crv in PerimeterCrvs)
            {
                SurfPoints.Add(crv.StartPoint);
            }

            if (!update) // Create new 
            {
                List<string> Profilepts = new List<string>();
                foreach (var v in SurfPoints)
                {
                    string dummy = null;
                    ret = Model.PointObj.AddCartesian(v.X * SF, v.Y * SF, v.Z * SF, ref dummy);
                    Profilepts.Add(dummy);
                }

                string[] names = Profilepts.ToArray();
                string dummyarea = string.Empty;
                ret = Model.AreaObj.AddByPoint(Profilepts.Count(), ref names, ref dummyarea);
                if (ret == 1) counter++;
                Id = dummyarea;
            }
            else
            {
                //update shell
            }
            if (counter > 0) Error = string.Format("Error creating mesh{0}", Id);
        }

        public static void CreateorUpdateJoint(ref cSapModel Model, Point pt, ref string Id, bool Update, Double SF)
        {
            if (!Update) //create new Joint 
            {
                string dummy = string.Empty;
                long ret = Model.PointObj.AddCartesian(pt.X * SF, pt.Y * SF, pt.Z * SF, ref dummy);
                Id = dummy;
                ret = Model.PointObj.SetSpecialPoint(Id, true);
            }
            else
            {
                //get coordiante and compre 
                double jx = 0;
                double jy = 0;
                double jz = 0;
                long ret = Model.PointObj.GetCoordCartesian(Id, ref jx, ref jy, ref jz);

                if (pt.X !=jx || pt.Y != jy || pt.Z != jz)
                {
                    string dummy = string.Empty;
                    ret = Model.PointObj.AddCartesian(pt.X, pt.Y, pt.Z, ref dummy);  // return exisiting data if exist 
                    if (dummy != Id)
                    {
                        //ret = Model.EditPoint.Change
                        //ret = Model.PointObj.ChangeName(dummy, Id); 
                    }
                }
            }
        }

        public static bool ChangeNameSAPfrm(ref cSapModel Model,string Name, string NewName)
        {
            long ret = Model.FrameObj.ChangeName(Name, NewName);
            if (ret == 0) { return true; } else { return false; } 
        }
        public static bool ChangeNameSAPArea(ref cSapModel Model, string Name, string NewName)
        {
            long ret = Model.AreaObj.ChangeName(Name, NewName);
            if (ret == 0) { return true; } else { return false; }
        }
        public static bool ChangeNameSAPJoint(ref cSapModel Model, string Name, string NewName)
        {
            long ret = Model.PointObj.ChangeName(Name, NewName);
            if (ret == 0) { return true; } else { return false; }
        }

        // to extract the Section Names on Specific Section Catalog
        public static void GetSectionsfromCatalog(ref cSapModel Model, string SC, ref string[] Names)
        {
            int number = 0;
            eFramePropType[] PropType = null;
            long ret = Model.PropFrame.GetPropFileNameList(SC, ref number, ref Names, ref PropType);
        }

    }
    
}
