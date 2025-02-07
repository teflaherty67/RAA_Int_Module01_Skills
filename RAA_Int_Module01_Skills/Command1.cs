
namespace RAA_Int_Module01_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // 00. create a transaction
            using(Transaction t = new Transaction(curDoc))
            {
                // start the transaction
                t.Start("Create Schedule");

                // 01. Create Schedule

                // get ElementId of desired category
                ElementId catId = new ElementId(BuiltInCategory.OST_Doors);

                // create the schedule & name it
                ViewSchedule newSchedule = ViewSchedule.CreateSchedule(curDoc, catId);
                newSchedule.Name = "My Door Schedule";

                // 02a. get parameters for fields

                // get all the doors
                FilteredElementCollector colDoors = new FilteredElementCollector(curDoc)
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .WhereElementIsNotElementType();

                // get first door from the list
                Element doorInst = colDoors.FirstElement();

                Parameter paramDrNum = doorInst.LookupParameter("Mark");
                Parameter paramDrLevel = doorInst.LookupParameter("Level");

                Parameter paramDrWidth = doorInst.get_Parameter(BuiltInParameter.DOOR_WIDTH);
                Parameter paramDrHeight = doorInst.get_Parameter(BuiltInParameter.DOOR_HEIGHT);
                Parameter paramDrType = doorInst.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM);

                // 02b. create the fields
                ScheduleField fieldDrNum = newSchedule.Definition.AddField(ScheduleFieldType.Instance, paramDrNum.Id);
                ScheduleField fieldDrLevel = newSchedule.Definition.AddField(ScheduleFieldType.Instance, paramDrLevel.Id);
                ScheduleField fieldDrWidth = newSchedule.Definition.AddField(ScheduleFieldType.ElementType, paramDrWidth.Id);
                ScheduleField fieldDrHeight = newSchedule.Definition.AddField(ScheduleFieldType.ElementType, paramDrHeight.Id);
                ScheduleField fieldDrType = newSchedule.Definition.AddField(ScheduleFieldType.Instance, paramDrType.Id);

                // set field properties
                fieldDrLevel.IsHidden = true;
                fieldDrWidth.DisplayType = ScheduleFieldDisplayType.Totals;

                // 03. filter by level
                Level filterLevel = GetLevelByName(curDoc, "01 - Entry Level");

                ScheduleFilter levelFilter = new ScheduleFilter(fieldDrLevel.FieldId, ScheduleFilterType.Equal, filterLevel.Id);
                newSchedule.Definition.AddFilter(levelFilter);

                // 04a. sort by door type
                ScheduleSortGroupField sortType = new ScheduleSortGroupField(fieldDrType.FieldId);
                sortType.ShowHeader = true;
                sortType.ShowFooter = true;
                sortType.ShowBlankLine = true;
                newSchedule.Definition.AddSortGroupField(sortType);

                // 04b. sort by door mark
                ScheduleSortGroupField sortMark = new ScheduleSortGroupField(fieldDrNum.FieldId);
                newSchedule.Definition.AddSortGroupField(sortMark);

                // 05. calculate totals
                newSchedule.Definition.IsItemized = true;
                newSchedule.Definition.ShowGrandTotal = true;
                newSchedule.Definition.ShowGrandTotalTitle = true;
                newSchedule.Definition.ShowGrandTotalCount = true;

                t.Commit();
            }

            // code snippit for filtering a list for unique items
            List<string> rawStrings = new List<string>() { "a", "a", "d", "c", "c", "d", "b", "d" };
            List<string> uniqueStrings = rawStrings.Distinct().ToList();
            uniqueStrings.Sort();

            return Result.Succeeded;
        }

        private Level GetLevelByName(Document curDoc, string levelName)
        {
            FilteredElementCollector m_colLevels = new FilteredElementCollector(curDoc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType();

            foreach (Level curLevel in m_colLevels)
            {
                if (curLevel.Name == levelName)
                    return curLevel;
            }

            return null;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData.Data;
        }
    }

}
