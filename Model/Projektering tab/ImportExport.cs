using Newtonsoft.Json;
using System.IO;

namespace Dalton_Trapper.Model.Projektering_tab
{
    public class ImportExport : Utilities.ViewModelBase
    {
        public static void SaveToJson(ProjectInfo data, string filePath)
        {
            var json = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }

        public static ProjectInfo LoadFromJson(string filePath)
        {
            var json = File.ReadAllText(filePath);
            return JsonConvert.DeserializeObject<ProjectInfo>(json);
        }
    }
}
