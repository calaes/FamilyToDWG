using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

[TransactionAttribute(TransactionMode.Manual)]
[RegenerationAttribute(RegenerationOption.Manual)]
public class FamilyToDWG : IExternalCommand
{
    public Result Execute(
        ExternalCommandData commandData,
        ref string message,
        ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;

        string keystr = @"FamilyToDWG/Settings";
        RegistryKey key = Registry.CurrentUser.OpenSubKey(keystr);
        if (key == null)
        {
            key = Registry.CurrentUser.CreateSubKey(keystr);
        }
        else
        {
            //string docdir = @"C:\Users\calaes\Documents\Visual Studio 2017\Projects\FamilyToDWG\FamilyToDWG\Mech Template.rte";
            var templateFD = new OpenFileDialog();
            templateFD.Filter = "rte files (*.rte)|*.rte";
            templateFD.ShowDialog();
            string docdir = templateFD.FileName;
            key.SetValue("TemplateLocation", @docdir);
        }
        UIDocument uiDoc = uiApp.OpenAndActivateDocument(key.GetValue("TemplateLocation").ToString());

        var FD = new OpenFileDialog();
        FD.Filter = "gfcj files (*.gfcj)|*.gfcj";
        FD.ShowDialog();

        string filename = FD.SafeFileName;
        string capsfiledir = FD.FileName.Replace(FD.SafeFileName, "") + FD.SafeFileName.Replace(".gfcj","");

        if (File.Exists(capsfiledir))
        {
            string[] famfiles = System.IO.Directory.GetFiles(capsfiledir, "*rfa");

            //List<ElementID> _added_element_ids = List<ElementId>();

            foreach (string famfile in famfiles)
            {
                using (Transaction tx = new Transaction(uiDoc.Document))
                {
                    tx.Start("Load Family");
                    Autodesk.Revit.DB.Family family = null;
                    uiDoc.Document.LoadFamily(famfile, out family);
                    tx.Commit();

                    string name = family.Name;
                    //ISet<ElementID> familySymbolIds = family.GetFamilySymbolIds();
                    foreach (ElementId id in family.GetFamilySymbolIds())
                    {
                        FamilySymbol famsymbol = family.Document.GetElement(id) as FamilySymbol;
                        XYZ origin = new XYZ(0, 0, 0);
                        tx.Start("Load Family Member");
                        famsymbol.Activate();
                        FamilyInstance instance = uiDoc.Document.Create.NewFamilyInstance(origin, famsymbol, StructuralType.NonStructural);
                        tx.Commit();

                        DWGExportOptions options = new DWGExportOptions();
                        //options.FileVersion = (ACADVersion)(3);

                        // Export the active view
                        ICollection<ElementId> views = new List<ElementId>();
                        views.Add(uiDoc.Document.ActiveView.Id);
                        string dwgfilename = famsymbol.Name + ".dwg";
                        uiDoc.Document.Export(@capsfiledir, @dwgfilename, views, options);

                        tx.Start("Delete Family Member");
                        uiDoc.Document.Delete(instance.Id);
                        tx.Commit();
                    }
                }
            }
        }
        else
        {
            Console.WriteLine("Please Create Export Directory For the chosen CAPS file.");
        }

        return Result.Succeeded;
    }
}
