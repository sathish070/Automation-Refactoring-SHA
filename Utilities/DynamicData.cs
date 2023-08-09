using Microsoft.VisualBasic.FileIO;

namespace SHAProject.Utilities
{
    public class LoginData
    {
        public string? Website { get; set; }
        public string? BrowserName { get; set; }
        public string? UserName { get; set; }
        public string? Password { get; set; }
        public string? loginFolderPath { get; set; }    

    }

    public class CurrentBrowser
    {
        public string? BrowserName { get; set; }
    }
    public enum WidgetCategories
    {
        XfCustomview = -2,
        XfStandardDose = -1,
        XfStandardBlank = 0,
        XfStandard = 1,
        XfAtp = 2,
        XfMst = 3,
        XfGra = 4,
        XfCellEnergy = 5,
        XfSubOx = 6,
        XfTCell = 7,
        XfTCellFitness = 8,
        XfTCellPersistence = 9,
        XfMitoDose = 10,
        XfMitoScreening = 11,
        XfAtpDose = 12,
        XfAtpScreening = 13,
        Common = 14
    }

    public enum WidgetTypes
    {
        // Initial Widget Types
        KineticGraphEcar = -1,
        KineticGraphPer = 0,
        AcuteResponse = 1,
        AtpProductionCoupledRespiration = 2,
        AtpProductionRateBasal = 3,
        AtpProductionRateData = 4,
        AtpProductionRateInduced = 5,
        BarChart = 6,
        CouplingEfficiencyPercent = 7,
        Basal = 8,
        BasalGlycolysis = 9,
        BaselineOcr = 10,
        BaselineEcar = 11,
        CompensatoryGlycolysis = 12,
        DataTable = 13,
        EnergeticMapBasal = 14,
        EnergeticMapInduced = 15,
        EnergyMap = 16,
        GlycoAtpProductionRate = 17,
        InducedGlycolysis = 18,
        KineticGraph = 19,
        MaximalRespiration = 20,
        MetabolicPotentialOcr = 21,
        MetabolicPotentialEcar = 22,
        MitoAtpProductionRate = 23,
        MitochondrialRespiration = 24,
        NonMitoO2Consumption = 25,
        PercentPerFromGlycolysisBasal = 26,
        PercentPerFromGlycolysisInduced = 27,
        ProtonLeak = 28,
        SpareRespiratoryCapacity = 29,
        SpareRespiratoryCapacityPercent = 30,
        StressedOcr = 31,
        StressedEcar = 32,
        XfAtpRateIndex = 33,
        XfCellEnergyPhenotype = 34,

        // XF T Cell
        GlycolyticActivity = 35,
        MaximalGlycolyticRate = 36,
        AreaUnderTheCurve = 37,

        // glycoPer
        GlycoPer = 38,

        // XF T Cell Persistence and Fitness
        PersistencePercent = 39,
        PercentAtpFromGlycolysisBasal = 40,
        PercentAtpFromGlycolysisInduced = 41,

        // XF Heatmap & Dose Response
        HeatMap = 42,
        DoseResponse = 43,

        // XF Mito Tox
        ZPrime = 44,
        ScreeningMap = 45,
        MitoToxIndex = 46,
        AtpSynthaseInhibition = 47,

        // XF ATP Dose & Screening
        TotalAtpProductionRate = 48,
        GlycoAtpProdRateHeatMap = 49,
        MitoAtpProdRateHeatMap = 50,
        TotalAtpProdRateHeatMap = 51,

        Comment = 52,
        WellImage = 53
    }

    public class ResultStatus
    {
        public bool Status { get; set; }
        public string? Message { get; set; }
    }

    public enum ChartType
    {
        Amchart,
        CanvasJS,
        Bar
    }

    public enum ScreenshotType
    {
        Info,
        Error
    }

    public class FileUploadOrExistingFileData
    {
        public string[]? PreRequest { get; set; }
        public bool IsFileUploadRequired { get; set; }
        public string? FileUploadPath { get; set; }
        public string? FileName { get; set; }
        public string? FileExtension { get; set; }
        public bool IsTitrationFile { get; set; }
        public bool IsNormalized { get; set; }
        public FileType FileType { get; set; }
        public bool OpenExistingFile { get; set; }
        public List<WidgetTypes>? SelectedWidgets { get; set; }
        public string? OligoInjection { get; set; }
    }

    public enum FileType
    {
        Xfe24 = 1,
        Xfe96 = 2,
        Xfp = 3,
        XfHsMini = 4,
        XFPro = 5
    }

    public enum PlateMapWell
    {
        Row24WellCount = 4,
        Row96or8Well = 8,
        Column8WellCount = 1,
        Column24WellCount = 6,
        Column96WellCount = 12
    }

    public class GraphSettings
    {
        public bool SynctoView { get; set; }
        public bool RemoveZoom { get; set; }
        public bool RemoveZeroLine { get; set; }
        public bool RemoveXAutoScale { get; set; }
        public bool RemoveYAutoScale { get; set; }
        public bool RemoveDataPointSymbols { get; set; }
        public bool RemoveRateHighlight { get; set; }
        public bool RemoveInjectionMarkers { get; set; }

        // Dose graph settings

        public bool DoseSynctoView { get; set; }
        public bool RemoveDoseZoom { get; set; }
        public bool RemoveDoseZeroLine { get; set; }
        public bool RemoveDoseXAutoScale { get; set; }
        public bool RemoveDoseYAutoScale { get; set; }
        public bool RemoveDoseDataPointSymbols { get; set; }
    }

    public class WidgetItems
    {
        public string? Measurement { get; set; }
        public string? Rate { get; set; }
        public string? Display { get; set; }
        public string? Y { get; set; }
        public bool Normalization { get; set; }
        public string? SortBy { get; set; }
        public string? NonBoxPlotFile { get; set; }
        public string? ErrorFormat { get; set; }
        public bool BackgroundCorrection { get; set; }
        public string? Baseline { get; set; }
        public string? ExpectedGraphUnits { get; set; }
        public bool GraphSettingsVerify { get; set; }
        public bool DataTableSettingsVerify { get; set; }
        public GraphSettings? GraphSettings { get; set; }
        public KitValidation? KitValidation { get; set; }
        public HeatTolerance? HeatTolerance { get; set; }
        public bool CheckNormalizationWithPlateMap { get; set; }
        public bool IsExportRequired { get; set; }
        public bool PlateMapSynctoView { get; set; }
        public bool DoseResponseAddWidget { get; set; }
        public bool DoseResponseAddView { get; set; }
        public string? Oligo { get; set; }
        public string? Induced { get; set; }
    }

    public class KitValidation
    {
        public bool AssayKitValidation { get; set; }
        public string? CatNumber { get; set; }
        public string? LotNumber { get; set; }
        public string? SWID { get; set; }
    }
    public class NormalizationData
    {
        public List<string>? Values { get; set; }
        public string? Units { get; set; } = null;
        public string? ScaleFactor { get; set; } = "1";
    }
    public class HeatTolerance
    {
        public bool ColourOptions { get; set; }
        public string? ColourTolerance { get; set; }
    }
    public static class PlateMapName
    {
        public static List<string> GetXfpWellName()
        {
            return new() { "A01", "B01", "C01", "D01", "E01", "F01", "G01", "H01" };
        }
        public static List<string> GetXfe24WellName()
        {
            return new() { "A01", "A02", "A03", "A04", "A05", "A06", "B01", "B02", "B03", "B04", "B05", "B06", "C01", "C02", "C03", "C04", "C05", "C06", "D01", "D02", "D03", "D04", "D05", "D06" };
        }
        public static List<string> GetXfe96WellName()
        {
            return new() { "A01", "A02", "A03", "A04", "A05", "A06", "A07", "A08", "A09", "A10", "A11", "A12", "B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12", "C01", "C02", "C03",
                    "C04", "C05", "C06", "C07", "C08", "C09", "C10", "C11", "C12", "D01", "D02", "D03", "D04", "D05", "D06", "D07", "D08", "D09", "D10", "D11", "D12", "E01", "E02", "E03", "E04", "E05", "E06", "E07", "E08", "E09", "E10",
                    "E11", "E12", "F01", "F02", "F03", "F04", "F05", "F06", "F07", "F08", "F09", "F10", "F11", "F12", "G01", "G02", "G03", "G04", "G05", "G06", "G07", "G08", "G09", "G10", "G11", "G12", "H01", "H02", "H03", "H04", "H05", "H06",
                    "H07", "H08", "H09", "H10", "H11", "H12" };
        }
    }

    public class FilesTabData
    {
        public bool LayoutVerification { get; set; }
        public bool PaginationVerification { get; set; }
        public string? PageNumber { get; set; }
        public string? FilesList { get; set; }
        public List<string>? searchBoxDataList { get; set; }
        public string? FileFirstName { get; set; }
        public string? FileMiddleName { get; set; }
        public string? FileLastName { get; set; }
        public string? FileFullName { get; set; }
        public string? Categories { get; set; }
        public string? Date { get; set; }
        public string? Instrument { get; set; }
        public string? License { get; set; }
        public bool CreateNewFolder { get; set; }
        public string? FolderName { get; set; }
        public string? SubFolderName { get; set; }
        public string? LastFolderName { get; set; }
        public string? FileUploadPath { get; set; }
        public string? FileName { get; set; }
        public string? FileExtension { get; set; }
        public string? FileLocatedFolderPath { get; set; }
        public string? AddCategories { get; set; }
        public string? Rename { get; set; }
        public bool DownloadFileVerification { get; set; }
        public bool MakeACopy { get; set; }
        public string? CopyFilePath { get; set; }
        public bool MoveToFolder { get; set; }
        public string? FolderPath { get; set; }
        public string? ReplaceOrRename { get; set; }
        public bool DeleteFile { get; set; }
        public bool AssayKitVerification { get; set; }
        public string? CatNumber { get; set; }
        public string? LotNumber { get; set; }
        public string? SWID { get; set; }
        public string? FolderRename { get; set; }
        public bool DeleteFolder { get; set; }
        public bool ExportFilesVerification { get; set; }
        public bool SendToVerfication { get; set; }
        public string? FirstMailRecepient { get; set; }
        public bool RenameVerification { get; set; }
        public string? FileRename { get; set; }
        public bool AddFavorite { get; set; }
        public bool AddCategory { get; set; }
        public string? AddCategoryName { get; set; }
        public string? EditCategoryName { get; set; }
    }

    public class WorkFlow5Data
    {
        public List<WidgetTypes> AddDoseWidget { get; set; }
        public bool AnalysisLayoutVerification { get; set; }
        public bool DeleteWidgetRequired { get; set; }
        public WidgetTypes DeleteWidgetName { get; set; }
        public bool AddWidgetRequired { get; set; }
        public WidgetTypes AddWidgetName { get; set; }
        public bool NormalizationVerification { get; set; }
        public bool ApplyToAllWidgets { get; set; }
        public bool ModifyAssay { get; set; }
        public WidgetTypes SelectWidgetName { get; set; }
        public bool GraphProperties { get; set; }
        public WidgetItems? KineticGraphOcr { get; set; }
        public WidgetItems? KineticGraphEcar { get; set; }
        public WidgetItems? KineticGraphPer { get; set; }
        public WidgetItems? Barchart { get; set; }
        public WidgetItems? EnergyMap { get; set; }
        public WidgetItems? HeatMap { get; set; }
        public WidgetItems? DoseResponseWidget { get; set; }
        public WidgetItems? DoseResponseView { get; set; }
        public WidgetItems? DoseResponse { get; set; }
        public bool CreateBlankView { get; set; }
        public WidgetTypes AddBlankWidget { get; set; }
        public string? CustomViewName { get; set; }
        public string? CustomViewDescription { get; set; }
        public string? AddGroupName { get; set; }
        public string? SelecttheControls { get; set; }
        public string? InjectionName { get; set; }
        public string NormalizedFileName { get; set; }
    }

    public class WorkFlow6Data
    {
        public bool AnalysisLayoutVerification { get; set; }
        public bool NormalizationVerification { get; set; }
        public string? NormalizationLabel { get; set; } = null;
        public string NormalizationScaleFactor { get; set; } = "1";
        public bool ApplyToAllWidgets { get; set; }
        public List<string> NormalizationValues { get; set; }
        public bool ModifyAssay { get; set; }
        public string? AddGroupName { get; set; }
        public string? SelecttheControls { get; set; }
        public string? InjectionName { get; set; }
        public WidgetItems MitochondrialRespiration { get; set; }
        public WidgetItems BasalRespiration { get; set; }
        public WidgetItems AcuteResponse { get; set; }
        public WidgetItems ProtonLeak { get; set; }
        public WidgetItems MaximalRespiration { get; set; }
        public WidgetItems SpareRespiratoryCapacity { get; set; }
        public WidgetItems NonmitoO2Consumption { get; set; }
        public WidgetItems ATPProductionCoupledRespiration { get; set; }
        public WidgetItems CouplingEfficiency { get; set; }
        public WidgetItems SpareRespiratoryCapacityPercentage { get; set; }
        public WidgetItems DataTable { get; set; }
        public string NormalizedFileName { get; set; }
    }

    public class WorkFlow7Data
    {
        public string[]? PreRequest2 { get; set; }
        public bool AnalysisLayoutVerification { get; set; }
        public string[]? PreRequest3 { get; set; }
        public bool Normalization { get; set; }
        public string NormalizationLabel { get; set; } = null;
        public string NormalizationScaleFactor { get; set; } = "1";
        public bool ApplyToAllWidgets { get; set; }
        public List<string>? NormalizationValues { get; set; }
        public string[]? PreRequest4 { get; set; }
        public bool ModifyAssay { get; set; }
        public string? ModifyAssayAddGroupName { get; set; }
        public string? ModifyAssaySelecttheControls { get; set; }
        public string? ModifyAssayInjectionName { get; set; }
        public WidgetItems? MitoATPProductionRate { get; set; }
        public WidgetItems? GlycoATPProductionRate { get; set; }
        public WidgetItems? ATPProductionRateData { get; set; }
        public WidgetItems? ATPProductionRate_Basal { get; set; }
        public WidgetItems? ATPproductionRate_Induced { get; set; }
        public WidgetItems? EnergeticMap_Basal { get; set; }
        public WidgetItems? EnergeticMap_Induced { get; set; }
        public WidgetItems? XFATPRateIndex { get; set; }
        public WidgetItems? DataTable { get; set; }
        public WidgetItems? Induced { get; set; }

        public string NormalizedFileName { get; set; }
    }
    public class WorkFlow8Data
    {
        public bool LayoutVerification { get; set; }
        public bool NormalizationVerification { get; set; }
        public WidgetTypes SelectWidgetName { get; set; }
        public bool GraphProperties { get; set; }
        public WidgetItems? CellEnergyPhenotype { get; set; }
        public WidgetItems? MetabolicPotentialOCR { get; set; }
        public WidgetItems? MetabolicPotentialECAR { get; set; }
        public WidgetItems? BaselineOCR { get; set; }
        public WidgetItems? BaselineECAR { get; set; }
        public WidgetItems? StressedOCR { get; set; }
        public WidgetItems? StressedECAR { get; set; }
        public WidgetItems? DataTable { get; set; }

    }
}