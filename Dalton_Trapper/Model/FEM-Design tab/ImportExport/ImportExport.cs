using Newtonsoft.Json;
using System.IO;

namespace Dalton_Trapper.Model.ImportExport
{
    public class ImportExport : Utilities.ViewModelBase
    {
        public static void SaveToJson(TransferData data, string filePath)
        {
            var json = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }

        public static TransferData LoadFromJson(string filePath)
        {
            var json = File.ReadAllText(filePath);
            return JsonConvert.DeserializeObject<TransferData>(json);
        }
    }
}
