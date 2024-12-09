namespace Dalton_Trapper.Model.ImportExport
{
    public class TransferData : Utilities.ViewModelBase
    {
        public List<Entries>? Coordinates { get; set; }
        public List<CurvedLine>? CurvedLines { get; set; }
        public List<Entries>? PointSupports { get; set; }
        public List<Entries>? LineSupports { get; set; }
        public List<Entries>? PointLoads { get; set; }
        public List<Entries>? LineLoads { get; set; }

        public string? Name { get; set; }
        public string? Thickness { get; set; }
        public string? Fck { get; set; }
        public string? CC { get; set; }
        public string? Q { get; set; }
        public string? Psi2q { get; set; }
        public string? ExtraG { get; set; }
        public string? Snow { get; set; }
        public string? Snow2 { get; set; }
        public bool? RunInBgYes { get; set; }
        public bool? RunInBgNo { get; set; }
        public bool? DisconnectYes { get; set; }
        public bool? DisconnectNo { get; set; }
        public bool? RunAnalysisYes { get; set; }
        public bool? RunAnalysisNo { get; set; }
        public bool? ApplyDocTemplateYes { get; set; }
        public bool? ApplyDocTemplateNo { get; set; }
        public bool? IsManual { get; set; }
        public string? SelectedEnvironmentClass { get; set; }
        public List<Reinforcement>? NetReinforcement { get; set; }
        public List<Reinforcement>? AdditionalReinforcement { get; set; }
    }
}
