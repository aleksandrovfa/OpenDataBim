using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Binding = Autodesk.Revit.DB.Binding;
using MessageBox = System.Windows.Forms.MessageBox;

namespace OpenDataBim
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            MessageBox.Show("Выбери csv файл c измененными данными");
            string filePath = "";
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.FileName;
            }
            else
            {
                MessageBox.Show("Файл не выбран");
                return Result.Succeeded;
            }
            

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Transaction tr = new Transaction(doc, "ChangeElements");
            tr.Start();

            using (var reader = new StreamReader(filePath))
            {
                string[] headers = reader.ReadLine().Split(',');

                while (!reader.EndOfStream)
                {

                    string[] values = reader.ReadLine().Split(',');

                    ElementId id = new ElementId(Convert.ToInt32(values[0]));
                    Element element = doc.GetElement(id);
                    if (element == null)
                    {
                        MessageBox.Show($"/n Элемент {values[0]} не найдет");
                        continue;
                    }


                    // Обходим все ячейки этой строки 
                    for (int i = 1; i < values.Length; i++)
                    {
                        if (values[i] == "")
                        {
                            continue;
                        }
                        if (headers[i] == "Type")
                        {
                            string nameSymbol = (doc.GetElement(element.GetTypeId()) as FamilySymbol).Name;
                            if (nameSymbol != values[i])
                            {
                                ChangeFamilyType(doc, element, values[i]);
                            }
                        }
                        if (headers[i] == "Level")
                        {
                            string nameLevel = doc.GetElement(element.LevelId).Name;
                            if (nameLevel != values[i])
                            {
                                ChangeLevel(doc, element, values[i]);
                            }
                        }
                        if (headers[i].StartsWith("_"))
                        {
                            if (element.LookupParameter(headers[i]) != null)
                            {
                                element.LookupParameter(headers[i]).Set(values[i]);
                            }
                            else
                            {
                                CreateProjectParameter(uiapp.Application, element, headers[i]);
                                element.LookupParameter(headers[i]).Set(values[i]);
                            }
                        }
                    }
                }
            }

            tr.Commit();

            return Result.Succeeded;
        }

        private bool ChangeLevel(Document doc, Element element, string levelName)
        {
            Level levelConnected = doc.GetElement(element.LevelId) as Level;
            try
            {
                Level levelForConnected = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(Level))
                .Cast<Level>()
                .Where(x => x.Name == levelName)
                .FirstOrDefault();
                if (levelForConnected == null) return false;


                double elevationBefore = levelConnected.Elevation + element
                    .get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                    .AsDouble();
                double elevationAfter = elevationBefore - levelForConnected.Elevation;
                element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(levelForConnected.Id);
                element.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(elevationAfter);
                return true;
            }
            catch (Exception ex)
            {
                string message = $"У элемента  {element.Id} не удалось изменить уровень c {levelConnected.Name} на {levelName}";
                message += "\r\n" + ex.ToString();
                MessageBox.Show(message);
                return false;
            }


        }

        private bool ChangeFamilyType(Document doc, Element element, string name)
        {
            string nameSymbol = (doc.GetElement(element.GetTypeId()) as FamilySymbol).Name;

            FamilySymbol familySymbol = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.Name == name)
                .FirstOrDefault();
            if (familySymbol == null)
            {
                string message = $"Не найден тип {name}";
                MessageBox.Show(message);
                return false;
            }

            try
            {
                familySymbol.Activate();
                element.ChangeTypeId(familySymbol.Id);
                return true;
            }
            catch (Exception ex)
            {
                string message = $"У элемента  {element.Id} не удалось изменить тип {nameSymbol} на {name}";
                message += "\r\n" + ex.ToString();
                MessageBox.Show(message);
                return false;
            }
        }

        public static void CreateProjectParameter(Application app, Element element, string name)
        {
            CategorySet cats = new CategorySet();
            cats.Insert(element.Category);

            string oriFile = app.SharedParametersFilename;
            string tempFile = Path.GetTempFileName() + ".txt";
            using (File.Create(tempFile)) { }
            app.SharedParametersFilename = tempFile;

            var defOptions = new ExternalDefinitionCreationOptions(name, ParameterType.Text)
            {
                Visible = true
            };
            ExternalDefinition def = app.OpenSharedParameterFile().Groups.Create("TemporaryDefintionGroup").Definitions.Create(defOptions) as ExternalDefinition;

            app.SharedParametersFilename = oriFile;
            File.Delete(tempFile);

            Binding binding = app.Create.NewTypeBinding(cats);
            binding = app.Create.NewInstanceBinding(cats);

            BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(def, binding, BuiltInParameterGroup.PG_TEXT);
        }

    }
}
