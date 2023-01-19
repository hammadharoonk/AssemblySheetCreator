using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace CustomAddin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AssemblySheetCooker : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                //Define a reference Object to accept the pick result


                //Pick a group
                Selection selection = uiapp.ActiveUIDocument.Selection;
                AssemblyPickFilter selFilter = new AssemblyPickFilter();
                List<Reference> pickedRefs = selection.PickObjects(ObjectType.Element, selFilter, "Please select relevant assemblies").ToList();
                List<AssemblyInstance> Assemblies = new List<AssemblyInstance>();
                List<ElementId> AssemblyTypeIds = new List<ElementId>();
                List<ElementId> AssemblyInstanceIds = new List<ElementId>();
                List<String> AssemblyTypeNames = new List<String>();

                foreach (Reference r in pickedRefs)
                {
                    Element e = doc.GetElement(r);
                    AssemblyInstance a = e as AssemblyInstance;
                    if (a.AssemblyTypeName.Contains("E") || a.AssemblyTypeName.Contains("I"))
                    {
                        Assemblies.Add(a);
                        AssemblyTypeIds.Add(a.GetTypeId());
                        AssemblyInstanceIds.Add(a.Id);
                        AssemblyTypeNames.Add(a.AssemblyTypeName);
                    }

                }
                string msg = "";
                TaskDialog.Show("Result", string.Format("This is updated Create Assembly Views"));

                var uniqueAssemblies = new HashSet<string>(AssemblyTypeNames);
                List<ElementId> UniqueAssemblyTypeIds = new List<ElementId>();
                List<ElementId> UniqueAssemblyInstanceIds = new List<ElementId>();
                List<int> uniqueIndices = new List<int>();

                foreach (string s in uniqueAssemblies)
                {
                    uniqueIndices.Add(AssemblyTypeNames.IndexOf(s));
                }

                for (int i = 0; i < uniqueIndices.Count; i++) UniqueAssemblyTypeIds.Add(AssemblyTypeIds[uniqueIndices[i]]);
                for (int i = 0; i < uniqueIndices.Count; i++) UniqueAssemblyInstanceIds.Add(AssemblyInstanceIds[uniqueIndices[i]]);
                for (int i = 0; i < uniqueIndices.Count; i++) msg += uniqueAssemblies.ToList()[i] + " " + UniqueAssemblyTypeIds[i].ToString() + "\n";


                //TaskDialog.Show("Result", string.Format("Number of assemblies selected: {0})\n  unique Assembly Types and TypeIds in selection:\n{1}", assemblies.Count, msg));



                ElementId ExtPlanViewId = GetViewTemplateIdByName("S_ASM_Panel_Ext_Assembly_Plan", doc);
                ElementId ExtBottomViewId = GetViewTemplateIdByName("S_ASM_Panel_Ext_Assembly_ElevationBottom", doc);
                ElementId ExtLeftViewId = GetViewTemplateIdByName("S_ASM_Panel_Ext_Assembly_ElevationLeft", doc);
                ElementId ExtRightViewId = GetViewTemplateIdByName("S_ASM_Panel_Ext_Assembly_ElevationRight", doc);
                ElementId ExtFrontViewId = GetViewTemplateIdByName("S_ASM_Panel_Ext_Assembly_ElevationFront", doc);

                ElementId IntPlanViewId = GetViewTemplateIdByName("S_ASM_Panel_Int_Assembly_Plan", doc);
                ElementId IntBottomViewId = GetViewTemplateIdByName("S_ASM_Panel_Int_Assembly_ElevationBottom", doc);
                ElementId IntLeftViewId = GetViewTemplateIdByName("S_ASM_Panel_Int_Assembly_ElevationLeft", doc);
                ElementId IntRightViewId = GetViewTemplateIdByName("S_ASM_Panel_Int_Assembly_ElevationRight", doc);
                ElementId IntFrontViewId = GetViewTemplateIdByName("S_ASM_Panel_Int_Assembly_ElevationFront", doc);

                ElementId TitleBlockTypeId = new ElementId(440638);

                XYZ planPt = new XYZ(0, 0.12, 0);
                XYZ bottomPt = new XYZ(0, -0.35, 0);
                XYZ leftPt = new XYZ(-0.23, -0.1, 0);
                XYZ rightPt = new XYZ(0.23, -0.1, 0);
                XYZ frontPt = new XYZ(0, -0.1, 0);

                bool yes = true;

                Transaction trans = new Transaction(doc);
                trans.Start("Lab");
                TaskDialog.Show("Result", string.Format("Started"));

                for (int i = 0; i < UniqueAssemblyInstanceIds.Count; i++)
                {
                    ElementId ass = UniqueAssemblyInstanceIds[i];
                    AssemblyInstance instanceAssembly = doc.GetElement(ass) as AssemblyInstance;
                    string nameAssembly = instanceAssembly.AssemblyTypeName;

                    ViewSection plan, left, right, front, bottom;
                    plan = left = right = front = bottom = null;
                    ViewSheet viewSheet = null;

                    if (nameAssembly.Contains('E'))
                    {
                        plan = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationTop, ExtPlanViewId, yes);
                        left = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationLeft, ExtLeftViewId, yes);
                        right = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationRight, ExtRightViewId, yes);
                        front = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationFront, ExtFrontViewId, yes);
                        bottom = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationBottom, ExtBottomViewId, yes);
                        viewSheet = AssemblyViewUtils.CreateSheet(doc, ass, TitleBlockTypeId);
                        string numberSheet = "001-0" + nameAssembly.Substring(nameAssembly.Length - 2);
                        viewSheet.Name = "DIŞ PANEL İMALAT PAFTASI";
                        viewSheet.SheetNumber = numberSheet;
                    }

                    else if (nameAssembly.Contains('I'))
                    {
                        plan = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationTop, IntPlanViewId, yes);
                        left = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationLeft, IntLeftViewId, yes);
                        right = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationRight, IntRightViewId, yes);
                        front = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationFront, IntFrontViewId, yes);
                        bottom = AssemblyViewUtils.CreateDetailSection(doc, ass, AssemblyDetailViewOrientation.ElevationBottom, IntBottomViewId, yes);
                        viewSheet = AssemblyViewUtils.CreateSheet(doc, ass, TitleBlockTypeId);
                        string numberSheet = "002-0" + nameAssembly.Substring(nameAssembly.Length - 2);
                        viewSheet.Name = "İÇ PANEL İMALAT PAFTASI";
                        viewSheet.SheetNumber = numberSheet;
                    }
                    Viewport.Create(doc, viewSheet.Id, plan.Id, planPt);
                    Viewport.Create(doc, viewSheet.Id, front.Id, frontPt);
                    Viewport.Create(doc, viewSheet.Id, bottom.Id, bottomPt);
                    Viewport.Create(doc, viewSheet.Id, left.Id, leftPt);
                    Viewport.Create(doc, viewSheet.Id, right.Id, rightPt);
                }

                TaskDialog.Show("Result", string.Format("Success"));

                trans.Commit();
                return Result.Succeeded;

            }

            //If the user right-clicks or presses Esc, handle the exception
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }


        }

        public ElementId GetViewTemplateIdByName(string name, Document doc)
        {
            IEnumerable<View> views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.Name.Equals(name));
            View template = views.FirstOrDefault();
            return template.Id;
        }

    }


    public class AssemblyPickFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            return (e.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Assemblies));
        }
        public bool AllowReference(Reference r, XYZ p)
        {
            return false;
        }
    }
}
