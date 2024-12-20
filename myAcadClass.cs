using acApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.IO;
using System.Drawing;
using Autodesk.AutoCAD.Colors;
using System.Threading;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Threading.Tasks;
using System.Globalization;
using Autodesk.AutoCAD.Internal;
using System.Xml.Linq;
using System.Diagnostics;
using System.Text;
using System.Diagnostics.Eventing.Reader;
using System.CodeDom;
using System.Runtime.InteropServices.ComTypes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Autodesk.AutoCAD.GraphicsInterface;
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.Interop.Common;

[ComImport, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface    IOleObject
{
    void SetClientSite(IntPtr pClientSite);
    void GetClientSite(out IntPtr ppClientSite);
    void SetHostNames([MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);
    void Close(int dwSaveOption);
    void SetMoniker(int dwWhichMoniker, IntPtr pmk);
    void GetMoniker(int dwAssign, int dwWhichMoniker, out IntPtr ppmk);
    void InitFromData(IntPtr pDataObject, bool fCreation, int dwReserved);
    void GetClipboardData(int dwReserved, out IntPtr ppDataObject);
    void DoVerb(int iVerb, IntPtr lpmsg, IntPtr pActiveSite, int lindex, IntPtr hwndParent, IntPtr lprcPosRect);
    void EnumVerbs(out IntPtr ppEnumOleVerb);
    void Update();
    void IsUpToDate();
    void GetUserClassID(out Guid pClsid);
    void GetUserType(int dwFormOfType, [MarshalAs(UnmanagedType.LPWStr)] out string pszUserType);
    void SetExtent(int dwDrawAspect, IntPtr psizel);
    void GetExtent(int dwDrawAspect, IntPtr psizel);
    void Advise(IntPtr pAdvSink, out int pdwConnection);
    void Unadvise(int dwConnection);
    void EnumAdvise(out IntPtr ppenumAdvise);
    void GetMiscStatus(int dwAspect, out int pdwStatus);
    void SetColorScheme(IntPtr pLogpal);
}

public static class GlobalVariables // Or put it in your existing class as a static member
{
    // generate unique id for this session and create the global variable
    private static string _uniqueId = Guid.NewGuid().ToString();
    private static bool _foundMatchInFile = false;
    private static int _matchesFound = 0;
    private static bool _foundBestCandidate = false;
    private static List<string> _potentialDrawingNames = new List<string>();
    private static List<(string name, double numberValue, string numberPart, string layer)> _detectedCandidates = new List<(string name, double numberValue, string numberPart, string layer)>();

    public static string UniqueId
    {
        get { return _uniqueId; }
        set { _uniqueId = value; }
    }
    public static bool FoundMatchInFile
    {
        get { return _foundMatchInFile; }
        set { _foundMatchInFile = value; }
    }
    public static int MatchesFound
    {
        get { return _matchesFound; }
        set { _matchesFound = value; }
    }
    public static bool FoundBestCandidate
    {
        get { return _foundBestCandidate; }
        set { _foundBestCandidate = value; }
    }
    public static List<string> PotentialDrawingNames
    {
        get { return _potentialDrawingNames; }
        set { _potentialDrawingNames = value; }
    }
    public static List<(string name, double numberValue, string numberPart, string layer)> DetectedCandidates
    {
        get { return _detectedCandidates; }
        set { _detectedCandidates = value; }
    }

    // More complex example with thread safety (important in multithreaded environments like AutoCAD)
    private static int _counter = 0;
    private static readonly object _counterLock = new object();

    public static int Counter
    {
        get
        {
            lock (_counterLock)
            {
                return _counter;
            }
        }
        set
        {
            lock (_counterLock)
            {
                _counter = value;
            }
        }
    }

    // Example of a one-time initialization
    private static bool _isInitialized = false;
    public static void Initialize()
    {
        if (!_isInitialized)
        {
            // Perform one-time initialization here
            // ...
            _isInitialized = true;
        }
    }
}
public class MyAcadLib
{
    [CommandMethod("RunBatchOperations")]
    public void RunBatchOperations()
    {
        // Access the active AutoCAD editor (command line)
        Editor ed = acApp.DocumentManager.MdiActiveDocument.Editor;
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

        string filename = Path.GetFileName(acDoc.Name);

        ed.WriteMessage($"___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|\n");
        ed.WriteMessage($"_|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|__\n");
        ed.WriteMessage($"___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|\n");
        ed.WriteMessage($"_|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|__\n");
        ed.WriteMessage($"___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|\n");
        ed.WriteMessage($"_|___|                                                                ..            _|___|__\n");
        ed.WriteMessage($"___|__ .-==+%@@@*====+##*:.  .:=+%@@@%+==-==#@@@@+==:-=+@@@*=-..     .#%:.          ___|___|\n");
        ed.WriteMessage($"_|___|      #@@@=:     =@@@-.    :%@@%=.    .@@@@-.     :#=:        .+@@*:.         _|___|__\n");
        ed.WriteMessage($"___|__      #@@@=:     .@@@@-.   .=@@@#:    .=@@@#:.    :+:.        -%@@@+.         ___|___|\n");
        ed.WriteMessage($"_|___|      #@@@=:     :@@@*-.     #@@@+.   :##@@@+.   .*:.        .++@@@%-.        _|___|__\n");
        ed.WriteMessage($"___|__      #@@@=:...:+@%*=:.      .%@@@-.  =-=@@@#-  .*-.         ==:+@@@#:.       ___|___|\n");
        ed.WriteMessage($"_|___|      #@@@=-----=@@*.         :@@@@:..*:.%@@@=. *=:         :=:. #@@@+.       _|___|__\n");
        ed.WriteMessage($"___|__      #@@@=:     .*@@@-.      .=@@@#:#-: .@@@@--+-.        .*=-::-@@@@=.      ___|___|\n");
        ed.WriteMessage($"_|___|      #@@@=:      :@@@@-.      .*@@@*=-.  :@@@%*-.        .=-:....=@@@%-.     _|___|__\n");
        ed.WriteMessage($"___|__      #@@@=:      :@@@%-.       .%@@%=.   .#@@@=:.       .-=:     .+@@@*:     ___|___|\n");
        ed.WriteMessage($"_|___|      #@@@=:     .%@@#-:.        :@@*-     -@@+:.        -#=.      .%@@@=.    _|___|__\n");
        ed.WriteMessage($"___|__ .==+*%%%%#+===+#%#=-:.           -#=.     .+#-.     .-*%%%%*=:.-=+*#%%%%*==..___|___|\n");
        ed.WriteMessage($"_|___|  ..................              ....       ..       .......... ............._|___|__\n");
        ed.WriteMessage($"___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|\n");
        ed.WriteMessage($"_|___|___|___|                                       |___|___|___|___|___|___|___|___|___|__\n");
        ed.WriteMessage($"___|___|___|__            RunBatchOperations:        __|___|___|___|___|___|___|___|___|__|__\n");
        ed.WriteMessage($"_|___|___|___|             {filename}         _|__|__|___|___|___|___|___|___|___|__\n");
        ed.WriteMessage($"___|___|___|__                                       __|___|___|___|___|___|___|___|___|___|__\n");
        ed.WriteMessage($"___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|_\n");
        ed.WriteMessage($"_|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|___|__\n");

        acDoc = null;
    }
    
    [CommandMethod("ZoomExtents")]
    public void ZoomExtents()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (acDoc != null)
        {
            Editor ed = acDoc.Editor;
            ed.Command("._ZOOM", "E");
        }
        acDoc = null;
    }

    [CommandMethod("ZoomAllExtents")]
    public void ZoomAllExtents()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (acDoc != null)
        {
            Editor ed = acDoc.Editor;
            Database db = acDoc.Database;

            // Temporarily disable background plotting to suppress printer dialog warnings
            object oldBackgroundPlot = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("BACKGROUNDPLOT");
            object oldCmddia = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDDIA");

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMDDIA", 0); // Disables command dialog windows

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    DBDictionary layouts = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                    foreach (DBDictionaryEntry layoutEntry in layouts)
                    {
                        Layout layout = trans.GetObject(layoutEntry.Value, OpenMode.ForRead) as Layout;

                        // Switch to the layout
                        if (layout.LayoutName != "Model")
                        {
                            acDoc.Editor.Command("_.LAYOUT", "Set", layout.LayoutName);
                        }

                        // Apply zoom extents
                        ed.Command("._ZOOM", "E");
                    }

                    trans.Commit();
                }
            }
            finally
            {
                // Restore system variables to their original values
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", oldBackgroundPlot);
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMDDIA", oldCmddia);
            }
            db = null;
        }
        acDoc = null;
    }

    [CommandMethod("SuppressWarnings")]
    public void SuppressWarnings()
    {
        var acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = acDoc.Editor;

        try
        {
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("FILEDIA", 0);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}");
        }
        acDoc.SendStringToExecute("CMDECHO 0 ", true, false, false);
        acDoc = null;
    }

    [CommandMethod("ActivateFirstSheet")]
    public void ActivateFirstSheet()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Database db = acDoc.Database;
        Editor ed = acDoc.Editor;

        string filePath = db.Filename;
        string parentDirectory = Path.GetDirectoryName(filePath);
        if (parentDirectory == null)
        {
            ed.WriteMessage("Parent directory not found.\n");
            return;
        }
        string parentLetter = Path.GetFileName(parentDirectory)?.Substring(0, 1).ToUpper();

        using (Transaction acTrans = acDoc.TransactionManager.StartTransaction())
        {
            DBDictionary layoutDict = acTrans.GetObject(acDoc.Database.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
            Layout firstLayout = null;
            int lowestTabOrder = int.MaxValue;

            // Find "model" layout if parentLetter is "A"
            if (parentLetter == "A")
            {
                foreach (DBDictionaryEntry layoutEntry in layoutDict)
                {
                    Layout layout = acTrans.GetObject(layoutEntry.Value, OpenMode.ForRead) as Layout;

                    if (layout.LayoutName.ToLower() == "model")
                    {
                        firstLayout = layout;
                        break;
                    }
                }

                if (firstLayout != null)
                {
                    ed.WriteMessage("\nActivating sheet: model");
                    LayoutManager.Current.CurrentLayout = firstLayout.LayoutName;
                }
                else
                {
                    ed.WriteMessage("\n'Model' sheet not found.");
                }
            }
            else
            {
                // Find the first layout (not "model") with the lowest tab order
                foreach (DBDictionaryEntry layoutEntry in layoutDict)
                {
                    Layout layout = acTrans.GetObject(layoutEntry.Value, OpenMode.ForRead) as Layout;

                    if (layout.LayoutName.ToLower() != "model" && layout.TabOrder < lowestTabOrder)
                    {
                        lowestTabOrder = layout.TabOrder;
                        firstLayout = layout;
                    }
                }

                if (firstLayout != null)
                {
                    ed.WriteMessage($"\nActivating sheet: {firstLayout.LayoutName}");
                    LayoutManager.Current.CurrentLayout = firstLayout.LayoutName;
                }
                else
                {
                    ed.WriteMessage("\nNo sheets found to activate.");
                }
            }

            acTrans.Commit();
        }
        acDoc = null;
        db = null;
    }

    [CommandMethod("SetSystemVars")]
    public void TurnOffGridMode()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;

        try
        {
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("GRIDMODE", 0);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}");
        }
        try
        {
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("SNAPMODE", 0);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}");
        }
        acDoc = null;
    }

    [CommandMethod("SetLayerProperties")]
    public void SetLayerProperties()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Database db = acDoc.Database;
        Editor ed = acDoc.Editor;

        using (Transaction acTrans = db.TransactionManager.StartTransaction())
        {
            try
            {
                LayerTable layerTable = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                if (layerTable.Has("0"))
                {
                    LayerTableRecord layer0 = acTrans.GetObject(layerTable["0"], OpenMode.ForWrite) as LayerTableRecord;
                    layer0.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);
                    db.Clayer = layerTable["0"];
                }

                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CECOLOR", "BYLAYER");
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CELTYPE", "BYLAYER");

                ed.WriteMessage("\nLayer Properties have been set.");
                acTrans.Commit();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }
        acDoc = null;
        db = null;
    }

    [CommandMethod("PurgeAnonymousBlocks")]
    public static void PurgeAnonymousBlocks()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Database db = acDoc.Database;
        Editor ed = acDoc.Editor;

        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
            try
            {
                BlockTable blockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectIdCollection blockIdsToPurge = new ObjectIdCollection();
                List<string> purgedBlockNames = new List<string>();

                foreach (ObjectId blockId in blockTable)
                {
                    BlockTableRecord btr = trans.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;

                    if (btr.IsAnonymous)
                    {
                        blockIdsToPurge.Add(blockId);
                        purgedBlockNames.Add(btr.Name);
                    }

                    if (btr.IsFromExternalReference)
                    {
                        if (btr.XrefStatus == XrefStatus.Unreferenced)
                        {
                            blockIdsToPurge.Add(blockId);
                            purgedBlockNames.Add(btr.Name);
                            ed.WriteMessage($"\nUnreferenced Xref '{btr.Name}' marked for removal.");
                        }
                    }
                }

                if (blockIdsToPurge.Count > 0)
                {
                    db.Purge(blockIdsToPurge);

                    foreach (ObjectId blockId in blockIdsToPurge)
                    {
                        if (!blockId.IsErased)
                        {
                            DBObject obj = trans.GetObject(blockId, OpenMode.ForWrite);
                            obj.Erase();
                            ed.WriteMessage($"\nErased Xref '{obj}'.");
                        }
                    }

                    int purgeCount = purgedBlockNames.Count;
                    ed.WriteMessage($"\n- {purgeCount} objects purged, including unreferenced xrefs and anonymous blocks.");
                }
                else
                {
                    ed.WriteMessage("\nNo anonymous blocks or unreferenced xrefs found for purging.");
                }
                if (blockIdsToPurge.Count > 0)
                {
                    db.Purge(blockIdsToPurge);
                    foreach (ObjectId blockId in blockIdsToPurge)
                    {
                        if (!blockId.IsErased)
                        {
                            ed.WriteMessage($"\nBlock '{blockId}' could not be erased.");
                        }
                    }
                }

                trans.Commit();
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nAn error occurred during the purge: " + ex.Message);
                trans.Abort();
            }
        }
        acDoc = null;
        db = null;
    }

    [CommandMethod("BindXrefs")]
    public void BindXrefs()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;
        Database db = acDoc.Database;
        int bindCount = 0;

        try
        {
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                ObjectIdCollection xrefsToBind = new ObjectIdCollection();

                bindCount = 0;
                foreach (ObjectId btrId in blockTable)
                {
                    BlockTableRecord btr = acTrans.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;

                    // Check if this block is an xref
                    if (btr.IsFromExternalReference && !btr.IsResolved)
                    {
                        xrefsToBind.Add(btrId);
                        bindCount++;
                    }
                }

                if (xrefsToBind.Count > 0)
                {
                    db.BindXrefs(xrefsToBind, true);
                }

                acTrans.Commit();
            }

            ed.WriteMessage($"\n- {bindCount} XRefs have been bound.");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}");
        }
        acDoc = null;
        db = null;
    }

    [CommandMethod("ExportToACAD2018")]
    public void ExportToACAD2018()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;
        Database db = acDoc.Database;
        DrawingNameExtractor extractor = new DrawingNameExtractor();
        StringBuilder logBuilder = new StringBuilder();

        string extractedName = "";

        string debugLogFilePath = @"C:/Users/mnewman/Desktop/exportDebugLog.txt";
        string filesToDeletePath = @"C:\mn\filesToDelete.txt";
        // logBuilder.AppendLine($"\n[GlobalVariables.UniqueId]: {GlobalVariables.UniqueId}");

        try
        {
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                string currentFilePath = acDoc.Name;
                string directory = Path.GetDirectoryName(currentFilePath);
                string filename = Path.GetFileNameWithoutExtension(currentFilePath);
                string newFilename = filename;

                Regex regex = new Regex(@"^(?<yearProject>\d{4}-\d{4})\s+(?<dwgName>[A-Za-z0-9\-.]+)\s*(?<descriptors>.*?)\.dwg$", RegexOptions.IgnoreCase);
                Match match = regex.Match(filename + ".dwg");
                // logBuilder.AppendLine($"\nSession: {GlobalVariables.UniqueId} | filename: {filename}");

                if (!match.Success)
                {
                    logBuilder.AppendLine($"\nnewFilename: {newFilename}");
                    logBuilder.AppendLine("Regex match failed.");
                    logBuilder.AppendLine($"Filename: {filename}");
                }
                else
                {
                    string yearProject = match.Groups["yearProject"].Value;
                    string dwgName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(match.Groups["dwgName"].Value.Replace("-", ""));
                    string descriptors = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(match.Groups["descriptors"].Value.Replace("-", " "));

                    newFilename = $"{yearProject} {dwgName} {descriptors}.dwg";

                    DBDictionary layoutDict = acTrans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    int sheetCount = 0;
                    foreach (DBDictionaryEntry layoutEntry in layoutDict)
                    {
                        Layout layout = acTrans.GetObject(layoutEntry.Value, OpenMode.ForRead) as Layout;
                        if (!layout.LayoutName.Equals("Model", StringComparison.InvariantCultureIgnoreCase) &&
                            layout.LayoutName.IndexOf("Layout", StringComparison.InvariantCultureIgnoreCase) < 0)
                        {
                            sheetCount++;
                        }
                    }

                    List<string> extractedNames = extractor.GetPotentialDrawingNames(acDoc);

                    /*
                    foreach (string name in GlobalVariables.PotentialDrawingNames)
                    {
                        logBuilder.AppendLine($"\nFound a potential drawing name: {name}");
                    }
                    */

                    if (GlobalVariables.PotentialDrawingNames == null || !GlobalVariables.PotentialDrawingNames.Any())
                    {
                        File.AppendAllText(debugLogFilePath, logBuilder.ToString());
                        return;
                    }

                    string drawingLetter = extractor.GetDrawingLetterFromFilename(acDoc);
                    string bestDrawingName = extractor.GetBestDrawingName(GlobalVariables.PotentialDrawingNames, drawingLetter, dwgName);
                    extractedName = bestDrawingName;
                    string directoryName = Path.GetFileName(directory);

                    // logBuilder.AppendLine($"<<<--- dwgName: {dwgName} <<<--- directoryName: {directoryName} <<<--- extractedName: {extractedName}");

                    if (extractedName != null && !directoryName.Trim().Equals("A", StringComparison.OrdinalIgnoreCase))
                    {
                        dwgName = $"{extractedName}";
                        logBuilder.AppendLine($"\nNew drawing name to be applied to filename: {dwgName}");
                    }

                    if (sheetCount > 1 && !dwgName.EndsWith("x", StringComparison.InvariantCultureIgnoreCase) && !directoryName.Trim().Equals("A", StringComparison.OrdinalIgnoreCase))
                    {
                        dwgName += "x";
                    }
                    else if (dwgName.EndsWith("X"))
                    {
                        dwgName = dwgName.Substring(0, dwgName.Length - 1) + "x";
                    }
                    newFilename = $"{yearProject} {dwgName} {descriptors}".Trim();
                }

                string normalizedCurrentFilePath = Path.GetFullPath(currentFilePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                File.AppendAllText(filesToDeletePath, $"{normalizedCurrentFilePath}\n");

                string acadFilePath = Path.Combine(directory, $"NEW-{newFilename}.dwg");

                int counter = 1;
                while (File.Exists(acadFilePath))
                {
                    logBuilder.AppendLine($"----------------------------------------------");
                    logBuilder.AppendLine($"---------- WARNING | DUPLICATE FILE ----------");
                    logBuilder.AppendLine($"--------------- {newFilename} ---------------");
                    logBuilder.AppendLine($"----------------------------------------------");
                    acadFilePath = Path.Combine(directory, $"NEW-{newFilename}_{counter}.dwg");
                    counter++;
                }

                db.SaveAs(acadFilePath, DwgVersion.AC1027);
                ed.WriteMessage($"\nSaved file as ACAD2018: {acadFilePath}");

                acTrans.Commit();
            }
        }
        catch (System.Exception ex)
        {
            acDoc.Editor.WriteMessage($"\nError: {ex.Message}");
        }
        try
        {
            File.AppendAllText(debugLogFilePath, logBuilder.ToString());
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError writing to log file: {ex.Message}");
        }
        acDoc = null;
        db = null;
    }


    [CommandMethod("CloseCurrentDocument")]
    public void CloseCurrentDocument()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;

        // Check if a document is open
        if (acDoc != null)
        {
            ed.WriteMessage("\nClosing the current document...");
            // acDoc.ClearUndo();
            acDoc.CloseAndDiscard();
            acDoc.Dispose();
        }
        else
        {
            ed.WriteMessage("\nNo active document to close.");
        }
    }

    // new commandmethod that does nothing except print a message with a unique id
    [CommandMethod("PrintMessageWithUniqueId")]
    public void PrintMessageWithUniqueId()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;
        string uniqueId = Guid.NewGuid().ToString();
        ed.WriteMessage($"\n[{uniqueId}] Hello from PrintMessageWithUniqueId");
    }

    [CommandMethod("DetectMatchingNameInFile")]
    public void DetectMatchingNameInFile()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;
        Database db = acDoc.Database;
        DrawingNameExtractor extractor = new DrawingNameExtractor();
        StringBuilder reportBuilder = new StringBuilder();
        string foundMatchLogPath = @"C:/Users/mnewman/Desktop/foundMatchLog.txt";
        // reportBuilder.AppendLine($"\nCurrent file: " + acDoc.Name);

        try
        {
            extractor.GetPotentialDrawingNames(acDoc);

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                string currentFilePath = acDoc.Name;
                string directory = Path.GetDirectoryName(currentFilePath);
                string filename = Path.GetFileNameWithoutExtension(currentFilePath);
                string newFilename = filename;
                string directoryName = Path.GetFileName(directory);

                Regex regex = new Regex(@"^(?<yearProject>\d{4}-\d{4}) (?<dwgName>[A-Za-z0-9\-.]+)\s?(?<descriptors>.*?)\.dwg$", RegexOptions.IgnoreCase);
                Match match = regex.Match(filename + ".dwg");
                string dwgName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(match.Groups["dwgName"].Value.Replace("-", ""));

                string drawingLetter = extractor.GetDrawingLetterFromFilename(acDoc);
                string bestDrawingName = extractor.GetBestDrawingName(GlobalVariables.PotentialDrawingNames, drawingLetter, dwgName);

                // for all objects in detectedcandidates, if they match dwgName both case insensitive, then we have a match

                if (GlobalVariables.DetectedCandidates.Any())
                {
                    foreach (var candidate in GlobalVariables.DetectedCandidates)
                    {
                        if (candidate.name.Equals(dwgName, StringComparison.OrdinalIgnoreCase))
                        {
                            GlobalVariables.FoundMatchInFile = true;
                            break;
                        }
                    }
                }
                else
                {
                    // reportBuilder.AppendLine($"============! | No detectable strings in file for {filename}");
                }
                if (!GlobalVariables.FoundMatchInFile)
                {
                    if (GlobalVariables.DetectedCandidates.Any())
                    {
                        var bestGuess = GlobalVariables.DetectedCandidates.First();
                        reportBuilder.AppendLine($"==! ATTENTION | No match for {filename} | Best guess: {bestGuess.name}");
                    }
                    else
                    {
                        reportBuilder.AppendLine($"==! ATTENTION | No match for {filename} | No candidates available");
                    }
                }

                acTrans.Commit();
            }
        }
        catch (System.Exception ex)
        {
            acDoc.Editor.WriteMessage($"No detected candidates found for {ex.Message}");
            reportBuilder.AppendLine($"No detected candidates found for {ex.Message}");
        }

        // Write accumulated log messages to the file
        try
        {
            File.AppendAllText(foundMatchLogPath, reportBuilder.ToString());
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"No detected candidates found for {ex.Message}");
        }
    }

    public class DrawingNameExtractor
    {
        public List<string> GetPotentialDrawingNames(Document acDoc)
        {
            GlobalVariables.PotentialDrawingNames.Clear();
            Database db = acDoc.Database;
            Editor ed = acDoc.Editor;

            StringBuilder logBuilder = new StringBuilder();
            string debugLogFilePath = @"C:/Users/mnewman/Desktop/debugLog3.txt";
            string currentFilePath = Path.GetFileName(acDoc.Name); // Get the filename without the full path

            try
            {
                if (acDoc == null || db == null)
                {
                    throw new ArgumentNullException("Document or Database is null.");
                }

                string drawingLetter = GetDrawingLetterFromFilename(acDoc);
                string drawingNamePattern = $@"^{drawingLetter}\d+(\.\d+)?$";

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt == null)
                    {
                        throw new InvalidOperationException("\nError: Unable to access the block table.");
                    }

                    // Simplified filter to select all objects
                    TypedValue[] filter = new TypedValue[]
                    {
                new TypedValue((int)DxfCode.Start, "*") // Select all objects
                    };

                    SelectionFilter selectionFilter = new SelectionFilter(filter);
                    PromptSelectionResult selectionResult = ed.SelectAll(selectionFilter);

                    // Debugging output to CSV files
                    string allDetectedFilePath = @"C:/Users/mnewman/Desktop/allDetected.csv";

                    using (StreamWriter sw = new StreamWriter(allDetectedFilePath, true))
                    {
                        // Write the header if the file is new
                        if (new FileInfo(allDetectedFilePath).Length == 0)
                        {
                            sw.WriteLine("FileName,TextType,TextContent,PotentialDrawingName,IsMatch,Position,Height,Rotation,Layer,TextStyleId");
                        }

                        if (selectionResult.Status == PromptStatus.OK)
                        {
                            SelectionSet selectionSet = selectionResult.Value;

                            foreach (SelectedObject selObj in selectionSet)
                            {
                                if (selObj != null)
                                {
                                    string textType = string.Empty;
                                    string textContent = string.Empty;
                                    string position = string.Empty;
                                    double height = 0;
                                    double rotation = 0;
                                    string layer = string.Empty;
                                    string textStyleId = string.Empty;

                                    if (selObj.ObjectId.ObjectClass.DxfName == "TEXT")
                                    {
                                        DBText dbText = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as DBText;
                                        textType = "DBText";
                                        textContent = dbText.TextString;
                                        position = dbText.Position.ToString();
                                        height = dbText.Height;
                                        rotation = dbText.Rotation;
                                        layer = dbText.Layer;
                                        textStyleId = dbText.TextStyleId.ToString();
                                    }
                                    else if (selObj.ObjectId.ObjectClass.DxfName == "MTEXT")
                                    {
                                        MText mText = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as MText;
                                        textType = "MText";
                                        textContent = mText.Text;
                                        position = mText.Location.ToString();
                                        height = mText.TextHeight;
                                        rotation = mText.Rotation;
                                        layer = mText.Layer;
                                        textStyleId = mText.TextStyleId.ToString();
                                    }

                                    // Discard textContent longer than 10 characters
                                    if (textContent.Length > 10)
                                    {
                                        continue;
                                    }

                                    bool isMatch = false;
                                    string potentialDrawingName = string.Empty;

                                    if (Regex.IsMatch(textContent, drawingNamePattern))
                                    {
                                        Match drawingNameMatch = Regex.Match(textContent, drawingNamePattern);
                                        potentialDrawingName = drawingNameMatch.Value;
                                        // logBuilder.AppendLine($"\nPotential Drawing Name: {potentialDrawingName}");
                                        GlobalVariables.PotentialDrawingNames.Add(potentialDrawingName);
                                        isMatch = true;
                                        GlobalVariables.FoundBestCandidate = true;
                                    }

                                    if (!string.IsNullOrEmpty(textContent))
                                    {
                                        // Write the log entry in CSV format with double quotes to handle commas and newlines
                                        sw.WriteLine($"\"{currentFilePath}\",\"{textType}\",\"{textContent}\",\"{potentialDrawingName}\",\"{isMatch}\",\"{position}\",\"{height}\",\"{rotation}\",\"{layer}\",\"{textStyleId}\"");
                                    }
                                }
                            }
                        }
                        else
                        {
                            sw.WriteLine("No objects found.");
                        }
                    }

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error in GetPotentialDrawingNames: {ex.Message}", ex);
            }

            // Debugging: Log the contents of potentialDrawingNames
            logBuilder.AppendLine($"Extracted {GlobalVariables.PotentialDrawingNames.Count} potential drawing names in {currentFilePath}.");
            
            foreach (var name in GlobalVariables.PotentialDrawingNames)
            {
                logBuilder.AppendLine($"GlobalVariables.PotentialDrawingNames: {name}\n");
            }
            

            // Write accumulated log messages to the file
            try
            {
                File.AppendAllText(debugLogFilePath, logBuilder.ToString());
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError writing to log file: {ex.Message}");
            }

            return GlobalVariables.PotentialDrawingNames;
        }

        public string GetDrawingLetterFromFilename(Document acDoc)
        {
            string activeSheetName = LayoutManager.Current.CurrentLayout;
            string currentFilePath = acDoc.Name;
            string drawingLetter = "";

            if (string.IsNullOrEmpty(currentFilePath) || string.IsNullOrEmpty(activeSheetName))
            {
                throw new InvalidOperationException("Unable to retrieve the file path or active sheet name.");
            }

            // Extract drawing letter from existing filename
            string filename = Path.GetFileName(currentFilePath);
            string filenameWithoutExtension = Path.GetFileNameWithoutExtension(filename);
            string drawingLetterPattern = @"(?:^\d{4}-\d+\s*|\s+)([A-Z])";
            Match match = Regex.Match(filenameWithoutExtension, drawingLetterPattern);
            if (!match.Success)
            {
                throw new InvalidOperationException("Unable to extract DrawingLetter from filename.");
            }
            drawingLetter = match.Groups[1].Value;

            return drawingLetter;
        }

        public string GetBestDrawingName(List<string> extractedNames, string drawingLetter, string existingCandidateName)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            Editor ed = acDoc.Editor;
            StringBuilder logBuilder = new StringBuilder();
            string debugLogFilePath = @"C:/Users/mnewman/Desktop/debugLog2.txt";
            // string logFilePath = @"C:/Users/mnewman/Desktop/getBestDrawingNameLog.txt";

            logBuilder.AppendLine($"--- GetBestDrawingName called for {Path.GetFileName(acDoc.Name)} at {DateTime.Now} ---");
            logBuilder.AppendLine($"Existing Candidate Name: {existingCandidateName}");
            logBuilder.AppendLine("Extracted Names:");
            foreach (var name in extractedNames)
            {
                logBuilder.AppendLine($"- {name}");
            }

            if (extractedNames == null || extractedNames.Count == 0)
            {
                logBuilder.AppendLine("Nothing found. Returning existing candidate name.");
                // Write accumulated log messages to the file
                try
                {
                    File.AppendAllText(debugLogFilePath, logBuilder.ToString());
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError writing to log file: {ex.Message}");
                }
                return existingCandidateName;
            }

            string pattern = $@"^([A-Za-z]?){drawingLetter}([A-Za-z]*)(?<number>-?\d+(\.\d+)?)([A-Za-z]*)$";
            GlobalVariables.DetectedCandidates.Clear();

            // Process extracted names
            foreach (string name in extractedNames)
            {
                string cleanedName = name.Replace(" ", "").Trim();
                logBuilder.AppendLine($"Processing extracted name: {cleanedName}");

                Match match = Regex.Match(cleanedName, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string numberPart = match.Groups["number"].Value.TrimStart('-');
                    if (double.TryParse(numberPart, out double numericValue))
                    {
                        string layer = GetLayerForTextContent(cleanedName, db);
                        GlobalVariables.DetectedCandidates.Add((name: cleanedName.Replace("-", ""), numberValue: numericValue, numberPart: numberPart, layer: layer));
                        // logBuilder.AppendLine($"Added candidate: {cleanedName.Replace("-", "")}, Numeric Value: {numericValue}, Layer: {layer}");
                    }
                }
                else
                {
                    logBuilder.AppendLine("No match found.");
                }
            }

            logBuilder.AppendLine("Detected Candidates in GetBestDrawingName function:");
            foreach (var candidate in GlobalVariables.DetectedCandidates)
            {
                logBuilder.AppendLine($"Candidate: {candidate.name}, Numeric Value: {candidate.numberValue}, Number Part: {candidate.numberPart}, Layer: {candidate.layer}");
            }

            // Parse existingCandidateName
            string existingCleanedName = existingCandidateName.Replace(" ", "").Trim();
            Match existingMatch = Regex.Match(existingCleanedName, pattern, RegexOptions.IgnoreCase);
            if (existingMatch.Success)
            {
                string existingNumberPart = existingMatch.Groups["number"].Value.TrimStart('-');
                if (double.TryParse(existingNumberPart, out double existingNumericValue))
                {
                    var existingCandidate = (name: existingCleanedName.Replace("-", ""), numberValue: existingNumericValue, numberPart: existingNumberPart, layer: "");
                    bool isExistingCandidateDetected = GlobalVariables.DetectedCandidates.Any(c => c.name == existingCandidate.name && c.numberPart == existingCandidate.numberPart);

                    if (isExistingCandidateDetected)
                    {
                        logBuilder.AppendLine("Existing candidate is among detected candidates. Returning existing candidate name.");
                        // Write accumulated log messages to the file
                        try
                        {
                            File.AppendAllText(debugLogFilePath, logBuilder.ToString());
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError writing to log file: {ex.Message}");
                        }
                        return existingCandidate.name;
                    }
                }
            }
            else
            {
                logBuilder.AppendLine("An existing candidate match was not found.");
            }

            // Reorder the global DetectedCandidates list
            if (GlobalVariables.DetectedCandidates.Count > 0)
            {
                // logBuilder.AppendLine("Re-ordering Candidates... SANITY CHECK");
                GlobalVariables.DetectedCandidates = GlobalVariables.DetectedCandidates
                    .OrderByDescending(c => c.layer == "TBLK" && c.name == existingCandidateName ? 1 : 0)
                    .ThenByDescending(c => c.layer == "TBLK" ? 1 : 0)
                    .ThenByDescending(c => Regex.IsMatch(c.numberPart, @"^\d+(\.\d+)?$") ? 1 : 0)
                    .ThenByDescending(c => c.numberPart.Length)
                    .ThenBy(c => c.numberValue)
                    .ToList();

                logBuilder.AppendLine("Re-ordered Candidates:");
                foreach (var candidate in GlobalVariables.DetectedCandidates)
                {
                    logBuilder.AppendLine($"Candidate: {candidate.name}, Numeric Value: {candidate.numberValue}, Number Part: {candidate.numberPart}, Layer: {candidate.layer}");
                }

                var bestCandidate = GlobalVariables.DetectedCandidates.First();
                logBuilder.AppendLine($"First candidate in the list: {bestCandidate.name}");
                // Write accumulated log messages to the file
                try
                {
                    File.AppendAllText(debugLogFilePath, logBuilder.ToString());
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError writing to log file: {ex.Message}");
                }
                return bestCandidate.name;
            }

            logBuilder.AppendLine("Returning existing candidate name.");
            // Write accumulated log messages to the file
            try
            {
                File.AppendAllText(debugLogFilePath, logBuilder.ToString());
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError writing to log file: {ex.Message}");
            }
            return existingCandidateName;
        }

        private string GetLayerForTextContent(string textContent, Database db)
        {
            if (string.IsNullOrEmpty(textContent))
            {
                throw new ArgumentNullException(nameof(textContent), "Text content cannot be null or empty.");
            }

            // Start a transaction to iterate over entities
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Access the Block Table
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the model space (or the desired block/table)
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId objId in btr)
                {
                    Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;

                    if (entity is DBText dbText && dbText.TextString == textContent)
                    {
                        return dbText.Layer;
                    }
                    else if (entity is MText mText && mText.Contents == textContent)
                    {
                        return mText.Layer;
                    }
                }
            }

            return null; // Text not found
        }

        private void CollectPotentialDrawingNames(string textContent, string drawingNamePattern, List<string> potentialDrawingNames, string drawingLetter)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;
            Database db = acDoc.Database;
            StringBuilder logBuilder = new StringBuilder();
            string debugLogFilePath = @"C:/Users/mnewman/Desktop/debugLog3.txt";
            logBuilder.AppendLine($"/nInitializing log file for CollectPotentialDrawingNames debugging...");

            try
            {
                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    using (StreamWriter sw = new StreamWriter(debugLogFilePath, true))
                    {
                        MatchCollection matches = Regex.Matches(textContent, drawingNamePattern);
                        logBuilder.AppendLine($"/nCollecting drawing names...");

                        foreach (Match match in matches)
                        {
                            string potentialDrawingName = match.Value.Trim();
                            potentialDrawingName = potentialDrawingName.Replace(" ", "").Replace("-", "");

                            if (potentialDrawingName.Length <= 10)
                            {
                                // Check if the potentialDrawingName matches the expected pattern after cleanup
                                // It should start with the drawing letter followed by numbers, dots, and optionally a trailing uppercase letter
                                string pattern = $@"^([A-Za-z]?){drawingLetter}([A-Za-z]*)(?<number>-?\d+(\.\d+)?)([A-Za-z]*)$";
                                if (Regex.IsMatch(potentialDrawingName, pattern))
                                {
                                    GlobalVariables.PotentialDrawingNames.Add(potentialDrawingName);
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in CollectPotentialDrawingNames: {ex.Message}");
            }
        }

        private void ProcessBlockReference(BlockReference blockRef, Transaction acTrans, string drawingNamePattern, List<string> potentialDrawingNames, string drawingLetter)
        {
            // Process attributes
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                AttributeReference attRef = acTrans.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef != null)
                {
                    string attTextContent = attRef.TextString;

                    // Search for the DrawingName pattern
                    CollectPotentialDrawingNames(attTextContent, drawingNamePattern, potentialDrawingNames, drawingLetter);
                }
            }

            // Process nested entities within the block reference
            BlockTableRecord blockDef = acTrans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            foreach (ObjectId entId in blockDef)
            {
                Entity ent = acTrans.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (ent is DBText dbText)
                {
                    string dbTextContent = dbText.TextString;

                    // Search for the DrawingName pattern
                    CollectPotentialDrawingNames(dbTextContent, drawingNamePattern, potentialDrawingNames, drawingLetter);
                }
                else if (ent is MText mText)
                {
                    string mTextContent = mText.Text;

                    // Search for the DrawingName pattern
                    CollectPotentialDrawingNames(mTextContent, drawingNamePattern, potentialDrawingNames, drawingLetter);
                }
                else if (ent is BlockReference nestedBlockRef)
                {
                    // Recursively process nested block references
                    ProcessBlockReference(nestedBlockRef, acTrans, drawingNamePattern, potentialDrawingNames, drawingLetter);
                }
            }
        }
    }

    [CommandMethod("TestSheetDetector")]
    public void TestSheetDetector()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;
        Database db = acDoc.Database;

        try
        {
            string resultFilePath = @"C:/Users/mnewman/Desktop/testResults.txt";
            if (File.Exists(resultFilePath))
            {
                File.Delete(resultFilePath);
            }

            using (StreamWriter sw = new StreamWriter(resultFilePath, true))
            {
                if (acDoc == null || db == null)
                {
                    sw.WriteLine("\nError: No active document or database.");
                    return;
                }

                string activeSheetName = LayoutManager.Current.CurrentLayout;
                string currentFilePath = acDoc.Name;

                if (string.IsNullOrEmpty(currentFilePath) || string.IsNullOrEmpty(activeSheetName))
                {
                    sw.WriteLine("\nError: Unable to retrieve the file path or active sheet name.");
                    return;
                }

                string filename = Path.GetFileName(currentFilePath);
                string filenameWithoutExtension = Path.GetFileNameWithoutExtension(filename);
                sw.WriteLine($"|xxx| Testing: {filename}");

                // Extract the DrawingLetter from the filename
                string drawingLetterPattern = @"(?:^\d{4}-\d+\s*|\s+)([A-Z])";
                Match match = Regex.Match(filenameWithoutExtension, drawingLetterPattern);

                if (!match.Success)
                {
                    sw.WriteLine("\nError: Unable to extract DrawingLetter from filename.");
                    return;
                }

                string drawingLetter = match.Groups[1].Value;
                // Define the regex pattern to match the DrawingLetter followed by numbers, spaces, and periods
                string drawingNamePattern = $@"{drawingLetter}[0-9\.\-\s]{{1,9}}[A-Z]?";

                string currentTab = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CTAB").ToString();
                // sw.WriteLine($"<<< Filename: {filename} <<<--- Active Sheet Name: {activeSheetName} <<<--- Current Layout (CTAB): {currentTab}");
                // sw.WriteLine($"Extracted DrawingLetter: {drawingLetter}");

                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt == null)
                    {
                        sw.WriteLine("\nError: Unable to access the block table.");
                        return;
                    }

                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "TEXT") // DXF code for DBText
                    };

                    LayoutManager lm = LayoutManager.Current;
                    Layout layout = acTrans.GetObject(lm.GetLayoutId(currentTab), OpenMode.ForRead) as Layout;
                    BlockTableRecord btr = acTrans.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    BlockTableRecord modelSpace = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                    LayerTable lt = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                    SelectionFilter selectionFilter = new SelectionFilter(filter);
                    PromptSelectionResult selectionResult = ed.SelectAll(selectionFilter);

                    if (selectionResult.Status == PromptStatus.OK)
                    {
                        SelectionSet selectionSet = selectionResult.Value;

                        foreach (SelectedObject selObj in selectionSet)
                        {
                            if (selObj != null)
                            {
                                DBText dbText = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as DBText;
                                string ssText = dbText.TextString;
                                sw.WriteLine($"DBText from SelectionSet: {ssText}");
                                // Search for the DrawingName pattern
                                if (Regex.IsMatch(ssText, drawingNamePattern))
                                {
                                    Match drawingNameMatch = Regex.Match(ssText, drawingNamePattern);
                                    string potentialDrawingName = drawingNameMatch.Value;
                                    sw.WriteLine($"PotentialDrawingName: {potentialDrawingName}");
                                }
                            }
                        }
                    }

                    acTrans.Commit();
                }

            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}");
        }
    }

    private void ProcessBlockReference(BlockReference blockRef, Transaction acTrans, StreamWriter sw, string drawingNamePattern)
    {
        // Process attributes
        foreach (ObjectId attId in blockRef.AttributeCollection)
        {
            AttributeReference attRef = acTrans.GetObject(attId, OpenMode.ForRead) as AttributeReference;
            if (attRef != null)
            {
                string attTextContent = attRef.TextString;
                sw.WriteLine($"ProcessBlockReference (Nested) Attribute Contents: {attTextContent}");

                if (attTextContent.Contains("%<"))
                {
                    sw.WriteLine($"Found Field Expression in AttributeReference: {attTextContent}");
                }

                // Search for the DrawingName pattern
                if (Regex.IsMatch(attTextContent, drawingNamePattern))
                {
                    Match drawingNameMatch = Regex.Match(attTextContent, drawingNamePattern);
                    string potentialDrawingName = drawingNameMatch.Value;
                    sw.WriteLine($"PotentialDrawingName: {potentialDrawingName}");
                }
            }
        }

        // Process nested entities within the block reference
        BlockTableRecord blockDef = acTrans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
        foreach (ObjectId entId in blockDef)
        {
            Entity ent = acTrans.GetObject(entId, OpenMode.ForRead) as Entity;
            if (ent == null) continue;

            if (ent is DBText dbText)
            {
                //sw.WriteLine($"DBText is Visible: {dbText.Visible}");
                // Log that a DBText entity was found
                sw.WriteLine($"Found DBText Entity with Handle (Nested): {dbText.Handle}");

                string dbTextContent = dbText.TextString;
                sw.WriteLine($"DBText Contents (Nested): {dbTextContent}");

                if (dbTextContent.Contains("%<"))
                {
                    sw.WriteLine($"Found Field Expression in DBText: {dbTextContent}");
                }

                // Search for the DrawingName pattern
                if (Regex.IsMatch(dbTextContent, drawingNamePattern))
                {
                    Match drawingNameMatch = Regex.Match(dbTextContent, drawingNamePattern);
                    string potentialDrawingName = drawingNameMatch.Value;
                    sw.WriteLine($"PotentialDrawingName: {potentialDrawingName}");
                }
            }
            else if (ent is MText mText)
            {
                // Use mText.Text to get the plain text without formatting codes
                string mTextContent = mText.Text;
                //sw.WriteLine($"MText Contents: {mTextContent}");

                if (mTextContent.Contains("%<"))
                {
                    sw.WriteLine($"Found Field Expression in MText: {mTextContent}");
                }

                // Search for the DrawingName pattern
                if (Regex.IsMatch(mTextContent, drawingNamePattern))
                {
                    Match drawingNameMatch = Regex.Match(mTextContent, drawingNamePattern);
                    string potentialDrawingName = drawingNameMatch.Value;
                    sw.WriteLine($"PotentialDrawingName: {potentialDrawingName}");
                }
            }
            else if (ent is AttributeDefinition attDef)
            {
                string attDefContent = attDef.TextString;
                //sw.WriteLine($"AttributeDefinition Contents: {attDefContent}");

                if (attDefContent.Contains("%<"))
                {
                    sw.WriteLine($"Found Field Expression in AttributeDefinition: {attDefContent}");
                }

                // Search for the DrawingName pattern
                if (Regex.IsMatch(attDefContent, drawingNamePattern))
                {
                    Match drawingNameMatch = Regex.Match(attDefContent, drawingNamePattern);
                    string potentialDrawingName = drawingNameMatch.Value;
                    sw.WriteLine($"PotentialDrawingName: {potentialDrawingName}");
                }
            }
            // Check for Dimension
            else if (ent is Dimension dimension)
            {
                string dimensionText = dimension.DimensionText;
                sw.WriteLine($"Dimension Text: {dimensionText}");
            }
            // Check for MLeader (multileader)
            else if (ent is MLeader mleader)
            {
                // Get the multileader text
                string mleaderText = mleader.MText?.Contents ?? "";
                sw.WriteLine($"MLeader Text: {mleaderText}");
            }
            // Check for Table
            else if (ent is Table table)
            {
                for (int row = 0; row < table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        string cellText = table.Cells[row, col].TextString;
                        sw.WriteLine($"Table Cell [{row},{col}] Text: {cellText}");
                    }
                }
            }
            else if (ent is BlockReference blockRef3)
            {
                // Process attributes
                foreach (ObjectId attId in blockRef3.AttributeCollection)
                {
                    AttributeReference attRef = acTrans.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (attRef != null)
                    {
                        string attTextContent = attRef.TextString;
                        //sw.WriteLine($"Attribute Contents: {attTextContent}");

                        if (attTextContent.Contains("%<"))
                        {
                            sw.WriteLine($"Found Field Expression in AttributeReference: {attTextContent}");
                        }

                        // Search for the DrawingName pattern
                        if (Regex.IsMatch(attTextContent, drawingNamePattern))
                        {
                            Match drawingNameMatch = Regex.Match(attTextContent, drawingNamePattern);
                            string potentialDrawingName = drawingNameMatch.Value;
                            sw.WriteLine($"PotentialDrawingName: {potentialDrawingName}");
                        }
                    }
                }

                // Optionally, process nested entities within the block reference
                ProcessBlockReference(blockRef, acTrans, sw, drawingNamePattern);
            }
        }
    }

    [CommandMethod("RenameSheetsByParentDir")]
    public void RenameSheetsByParentDir()
    {
        Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        string filePath = db.Filename;
        string parentDirectory = Path.GetDirectoryName(filePath);
        if (parentDirectory == null)
        {
            ed.WriteMessage("Parent directory not found.\n");
            return;
        }
        string parentLetter = Path.GetFileName(parentDirectory)?.Substring(0, 1).ToUpper();

        // Define a regular expression that matches the pattern "Letter + alphanumeric with allowed special chars (dash, period, space)"
        Regex regex = new Regex($@"^{parentLetter}[A-Za-z0-9.\-\s]{{1,6}}$");

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                Layout layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                string newSheetName = null;

                foreach (ObjectId objId in btr)
                {
                    DBText dbText = tr.GetObject(objId, OpenMode.ForRead) as DBText;
                    MText mText = tr.GetObject(objId, OpenMode.ForRead) as MText;

                    string text = null;

                    if (dbText != null)
                    {
                        text = dbText.TextString;
                    }
                    else if (mText != null)
                    {
                        text = mText.Text;
                    }

                    if (text != null && regex.IsMatch(text))
                    {
                        string potentialSheetName = text.Replace("-", "").Replace(" ", "");

                        if (newSheetName == null)
                        {
                            newSheetName = potentialSheetName;
                            break;  // break on first valid match
                        }
                    }
                }

                if (!string.IsNullOrEmpty(newSheetName))
                {
                    layout.UpgradeOpen();
                    layout.LayoutName = newSheetName;
                    layout.DowngradeOpen();
                    ed.WriteMessage($"Renamed layout '{entry.Key}' to '{newSheetName}'.\n");
                }
            }

            tr.Commit();
        }
    }

    [CommandMethod("SearchSystemVariables")]
    public void SearchSystemVariables()
    {
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = acDoc.Editor;
        Database db = acDoc.Database;

        string resultFilePath = @"C:/Users/mnewman/Desktop/testResults.txt";

        if (File.Exists(resultFilePath))
        {
            File.Delete(resultFilePath);
        }

        using (StreamWriter sw = new StreamWriter(resultFilePath, true))  // Append mode
        {
            try
            {
                // Check some common system variables
                string currentTab = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CTAB").ToString();
                sw.WriteLine($"CTAB (Current Layout): {currentTab}");

                string fileName = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("DWGNAME").ToString();
                sw.WriteLine($"DWGNAME (Current Drawing Name): {fileName}");

                string filePath = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("DWGPREFIX").ToString();
                sw.WriteLine($"DWGPREFIX (Drawing Path): {filePath}");

                string userName = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LOGINNAME").ToString();
                sw.WriteLine($"LOGINNAME (User Name): {userName}");

                string sysVarValue = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("PLOTID").ToString();
                sw.WriteLine($"PLOTID: {sysVarValue}");

                // Check other system variables based on what might be relevant
                string acadVersion = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("ACADVER").ToString();
                sw.WriteLine($"ACADVER (AutoCAD Version): {acadVersion}");

                // You can add more system variables to check based on what might be affecting the text
            }
            catch (System.Exception ex)
            {
                sw.WriteLine($"\nError: {ex.Message}");
            }
        }

        ed.WriteMessage("\nSystem variables have been written to the result file.");
    }

}
