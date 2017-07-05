﻿/// Developed by Thornton Tomasetti's CORE Studio for Autodesk
/// http://core.thorntontomasetti.com
/// CORE Developers: Elcin Ertugrul and Ana Garcia Puyol

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;

namespace DynamoETABS.Definitions
{
    public class Restraint
    {
        //FIELDS

        internal bool u1;
        internal bool u2;
        internal bool u3;
        internal bool r1;
        internal bool r2;
        internal bool r3;

        // PUBLIC METHODS
        /// <summary>
        /// Set a Restraint on a node
        /// </summary>
        /// <param name="Point">Point to set the Restraint on</param>
        /// <param name="Tx">Translation X</param>
        /// <param name="Ty">Translation Y</param>
        /// <param name="Tz">Translation Z</param>
        /// <param name="Rx">Rotation X</param>
        /// <param name="Ry">Rotation Y</param>
        /// <param name="Rz">Rotation Z</param>
        /// <returns>Restraint</returns>
        public static Restraint Define(bool Tx, bool Ty, bool Tz, bool Rx, bool Ry, bool Rz)
        {
            return new Restraint(Tx, Ty, Tz, Rx, Ry, Rz);
        }

        // FAST RESTRAINTS
        /// <summary>
        /// Create a Fixed node
        /// </summary>
        /// <param name="Point">Point to set a Fixed Restraint</param>
        /// <returns>Fixed Restraint</returns>
        public static Restraint Fixed()
        {
            return new Restraint(true, true, true, true, true, true);
        }

        /// <summary>
        /// Create a Pinned node
        /// </summary>
        /// <param name="Point">Point to set a Pinned Restraint</param>
        /// <returns>Pinned Restraint</returns>
        public static Restraint Pinned()
        {
            return new Restraint(true, true, true, false, false, false);
        }
        /// <summary>
        /// Create a Roller
        /// </summary>
        /// <param name="Point">Point to set a Roller</param>
        /// <returns>Roller</returns>
        public static Restraint Roller()
        {
            return new Restraint(false, false, true, false, false, false);
        }
        /// <summary>
        /// Create a Simple Restraint
        /// </summary>
        /// <param name="Point">Point to set a simple restraint</param>
        /// <returns>Simple Restraint</returns>
        public static Restraint Simple()
        {
            return new Restraint(false, false, false, false, false, false);
        }

        /// <summary>
        /// Decompose a Restraint
        /// </summary>
        /// <param name="restraint">Restraint to decompose</param>
        /// <returns>Node point, U1, U2, U3, R1, R2 and R3 </returns>
        [MultiReturn("Tx", "Ty", "Tz", "Rx", "Ry", "Rz")]
        public static Dictionary<string, object> Decompose(Restraint restraint)
        {
            // Return outputs
            return new Dictionary<string, object>
            {
                {"Tx", restraint.u1},
                {"Ty", restraint.u2},
                {"Tz", restraint.u3},
                {"Rx", restraint.r1},
                {"Ry", restraint.r2},
                {"Rz", restraint.r3}
            };
        }

        // PRIVATE CONSTRUCTOR
        private Restraint() { }
        private Restraint(bool U1, bool U2, bool U3, bool R1, bool R2, bool R3)
        {
            u1 = U1;
            u2 = U2;
            u3 = U3;
            r1 = R1;
            r2 = R2;
            r3 = R3;
        }


    }
}
