using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;

namespace ML
{
    public class Class1
    {
        [CommandMethod("ML")]
        public void ML()
        {
            //Document acDoc = Application.DocumentManager.MdiActiveDocument;
            //Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\n专属定制工具(醉月武编写——QQ交流群：325959696)");

            FormML form = new FormML();
            Application.ShowModalDialog(form);
        }        
    }
}
