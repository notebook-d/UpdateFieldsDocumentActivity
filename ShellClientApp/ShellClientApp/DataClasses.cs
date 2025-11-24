using ProtoBuf;
using System.Runtime.InteropServices;

namespace ShellClientApp
{
    [ComVisible(false)]
    public enum ShellResult
    {
        Ok,
        Error
    }

    [ComVisible(false)]
    [ProtoContract]
    public class CommandInvokeData
    {
        [ProtoMember(1)]
        public Guid CommandId { get; set; }

        [ProtoMember(2)]
        public List<string> Paths { get; private set; }

        [ProtoMember(3)]
        public string ExtraArgs { get; set; }

        public CommandInvokeData()
        {
            Paths = new List<string>();
        }
    }

    [ComVisible(false)]
    [ProtoContract]
    public class CommandInvokeResult
    {
        [ProtoMember(1)]
        public ShellResult Result { get; set; }

        [ProtoMember(2)]
        public byte[] Data { get; set; }
    }

        [ComVisible(false)]
    public static class CommandIds
    {
        public static readonly Guid DescriptionNonInteractive = new Guid("F717AD8D-3435-4338-92B0-F256172D8BBB");
        public static readonly Guid Commit = new Guid("8458614C-1FE3-4DFA-B2BF-BB2AF8BEF320");
        public static readonly Guid Publish = new Guid("D7093F13-12B2-45B7-99A7-47CC886A38F9");
        public static readonly Guid PublishAndCommit = new Guid("F27B35B4-42BC-49AB-884A-C6C27AEE9F7A");
        public static readonly Guid UpdateFilesAttributes = new Guid("20A69A1E-AA18-430A-AB88-E30934B119D1");
        public static readonly Guid AppendPublish = new Guid("3428D3C4-37EE-462D-A5E6-8C6E23C7290F");
        public static readonly Guid AppendPublishAndCommit = new Guid("56078AB2-BEFC-44DC-9AF6-52AF0516FAC4");
        public static readonly Guid Discard = new Guid("EA1B8F98-AFDC-4328-9C35-B878AD9BAE8B");
        public static readonly Guid Download = new Guid("F0A2A4E1-458B-4B4C-A259-73DC69894278");
        public static readonly Guid GetLatestVersion = new Guid("FE8B7DF7-B785-47D6-950E-121D6D35509C");
        public static readonly Guid Revert = new Guid("246ABBE1-F7AC-4241-B4F4-47E3B95952CE");
        public static readonly Guid SharingSettings = new Guid("4F42CDEA-E3C6-4F95-B731-504CA71829FC");
        public static readonly Guid ShowDocument = new Guid("E3DA0178-03E7-4903-BE67-C2EA73E4753F");
        public static readonly Guid QuitClient = new Guid("966ecbe2-4b5f-464b-9e96-12bc10e93885");
        public static readonly Guid ShowAllPrinters = new Guid("ca895788-2238-481f-8ed4-406d109777bb");
        public static readonly Guid ShowHistory = new Guid("798B2FF1-80F0-4BAD-8C4E-4562466A99AA");
        public static readonly Guid ShowProjectsExplorer = new Guid("12940BB8-E819-43D8-8B1F-9F018C25BA5F");
        public static readonly Guid ShowRecycleBin = new Guid("C47B1458-CE84-407C-B7F6-A60C402275E0");
        public static readonly Guid Subsribe = new Guid("8271C139-A06F-4E51-B5C6-5200E90F6FA8");
        public static readonly Guid Unsubscribe = new Guid("2D559484-9A3A-4374-8C90-E9752DDE5A14");
        public static readonly Guid Unmount = new Guid("5EC20BE2-B47A-4644-8537-FF6A9AAB9517");
        public static readonly Guid BuildContextMenu = new Guid("54808C44-0091-4638-829B-FC574798464B");
        public static readonly Guid Lock = new Guid("BFB8A2C3-B460-43CD-BE97-D97E860CAE72");
        public static readonly Guid Unlock = new Guid("E6F5DE56-E8BE-490C-B79A-CCA6DCFDE380");
        public static readonly Guid AttachTo = new Guid("E559EB9D-19BA-4A63-8DD6-985CDE1EF399");
        public static readonly Guid AttachToTask = new Guid("56E9F934-1C0D-4F4B-BF8F-8FEA46CA2F21");
        public static readonly Guid AttachToReviewTask = new Guid("3F959427-4626-4508-9319-E0E148A12EDA");
        public static readonly Guid AttachToTaskFromTemplate = new Guid("928C714C-3329-496E-8C7E-69CDACD67676");
        public static readonly Guid CopyLink = new Guid("192E13FA-4450-4EFF-9578-1941668DC9CE");
        public static readonly Guid Freeze = new Guid("1C91C2B1-4E50-4B3C-9D13-8E8649EC2397");
        public static readonly Guid Unfreeze = new Guid("2D0B5E2E-41B8-464D-AB61-25BC9311083E");
        public static readonly Guid ShowProperties = new Guid("621137DD-90E9-41F9-B297-DF9304F448A3");
        public static readonly Guid GcCollect = new Guid("4B8334FD-13F0-4556-A068-B9402CFF95FC");
        public static readonly Guid ShowChat = new Guid("A66434E7-AD99-4AE8-970F-0A7C43801C90");
        public static readonly Guid AttachToTaskSubMenu = new Guid("EF92DC0A-D9C4-4068-A3FB-38C58C3744A1");
        public static readonly Guid AttachToWorkflowSubMenu = new Guid("F61BBEED-4C24-4D38-BEAB-9ACEEF83B318");
        public static readonly Guid Separator = new Guid("A2F3C6A7-CFFD-46E5-A5C3-10502982306F");
        public static readonly Guid ShowConnectionSettings = new Guid("BE127E4D-7FD1-4D89-B705-D8CD4C5A80D9");
    }
}
