#if UNITY_EDITOR
using UnityEditor;
using UnityEngine;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections.Generic;
using System.Security;
using System.Text;
using CompressionLevel = System.IO.Compression.CompressionLevel;

public class ScriptCollectorToWord : EditorWindow
{
    private const float LERP_SPEED = 10f;
    
    private string _rootFolder     = NormalizePath(Application.dataPath);
    private string _outputFolder   = NormalizePath(Application.dataPath);
    private string _outputFileName = "ScriptsBundle.docx";

    private bool   _showAdvanced         = false;
    private bool   _includeAllSubfolders = true;
    private bool   _excludeEditorFolders = false;

    private string[] _subfolders        = Array.Empty<string>();
    private bool[]   _subfolderSelected = Array.Empty<bool>();
    private Vector2  _subfoldersScroll  = Vector2.zero;

    private bool     _chooseConcreteScripts = false;
    private string[] _scriptFiles    = Array.Empty<string>(); 
    private bool[]   _scriptSelected = Array.Empty<bool>();
    private Vector2  _scriptsScroll  = Vector2.zero;

    private bool   _generationSucceeded = false;
    private string _generatedPath       = string.Empty;

    private float       _pendingHeight  = -1f;
    private static bool _windowCentered = false;

    [MenuItem("Tools/Generate Word from Scripts", priority = 250)]
    private static void OpenWindow()
    {
        var w = GetWindow<ScriptCollectorToWord>("Scripts → Word");
        w.minSize = new Vector2(580, 300);

        if (!_windowCentered)
        {
            var res = UnityStats.screenRes.Split('x');
            if (res.Length == 2 &&
                int.TryParse(res[0], out int sw) &&
                int.TryParse(res[1], out int sh))
            {
                var r = w.position;
                r.x = (sw - r.width)  * 0.5f;
                r.y = (sh - r.height) * 0.5f;
                w.position = r;
            }
            _windowCentered = true;
        }

        w.RefreshSubfolders();
        w.RefreshScripts();
        w.Repaint();
    }

    private void OnEnable()
    {
        RefreshSubfolders();
        RefreshScripts();
    }

    private void Update()
    {
        if (_pendingHeight < 0f) return;

        var r = position;
        r.height = Mathf.Lerp(r.height, _pendingHeight, Time.deltaTime * LERP_SPEED);

        if (Mathf.Abs(r.height - _pendingHeight) < 0.5f)
        {
            r.height      = _pendingHeight;
            _pendingHeight = -1f;
        }
        position = r;
    }

    private void OnGUI()
    {
        void Invalidate()          => _generationSucceeded = false;
        void FiltersChanged()      { RefreshScripts(); Invalidate(); }

        EditorGUILayout.Space(6);
        DrawRootFolderSelector(FiltersChanged);

        bool prevAdv = _showAdvanced;
        _showAdvanced = EditorGUILayout.ToggleLeft("Advanced options", _showAdvanced, EditorStyles.boldLabel);
        if (_showAdvanced != prevAdv) Invalidate();

        if (_showAdvanced)
        {
            EditorGUILayout.Space(4);
            DrawAdvancedPanel(FiltersChanged, Invalidate);
        }

        EditorGUILayout.Space(6);
        DrawOutputSelector(Invalidate);

        EditorGUILayout.Space(10);
        DrawSuccessBox();

        GUI.enabled = !string.IsNullOrWhiteSpace(_rootFolder) &&
                      !string.IsNullOrWhiteSpace(_outputFolder);

        if (GUILayout.Button("Generate Word file", GUILayout.Height(32)))
            GenerateDocx();

        GUI.enabled = true;

        if (Event.current.type == EventType.Repaint)
        {
            float need = GUILayoutUtility.GetLastRect().yMax + 10f;
            need = Mathf.Max(need, minSize.y);
            if (Mathf.Abs(position.height - need) > 0.5f)
                _pendingHeight = need;
        }
    }

    private void DrawRootFolderSelector(Action filtersChanged)
    {
        EditorGUILayout.LabelField("Root folder with scripts", EditorStyles.boldLabel);
        EditorGUILayout.BeginHorizontal();
        string newPath = EditorGUILayout.TextField(_rootFolder);
        if (newPath != _rootFolder)
        {
            _rootFolder = NormalizePath(newPath);
            RefreshSubfolders();
            filtersChanged();
        }
        if (GUILayout.Button("…", GUILayout.Width(28)))
        {
            string sel = EditorUtility.OpenFolderPanel("Select root folder", _rootFolder, "");
            if (!string.IsNullOrEmpty(sel))
            {
                _rootFolder = NormalizePath(sel);
                RefreshSubfolders();
                filtersChanged();
            }
        }
        EditorGUILayout.EndHorizontal();
    }

    private void DrawAdvancedPanel(Action filtersChanged, Action invalidate)
    {
        bool incAll = EditorGUILayout.ToggleLeft("Include all sub-folders", _includeAllSubfolders);
        if (incAll != _includeAllSubfolders) { _includeAllSubfolders = incAll; filtersChanged(); }

        bool exclEd = EditorGUILayout.ToggleLeft("Exclude folders named ‘Editor’", _excludeEditorFolders);
        if (exclEd != _excludeEditorFolders) { _excludeEditorFolders = exclEd; filtersChanged(); }

        if (!_includeAllSubfolders)
        {
            EditorGUILayout.LabelField("Select sub-folders to scan", EditorStyles.boldLabel);
            EditorGUILayout.BeginVertical(EditorStyles.helpBox);
            _subfoldersScroll = EditorGUILayout.BeginScrollView(_subfoldersScroll, GUILayout.Height(120));
            for (int i = 0; i < _subfolders.Length; i++)
            {
                bool sel = EditorGUILayout.ToggleLeft(_subfolders[i], _subfolderSelected[i]);
                if (sel != _subfolderSelected[i])
                {
                    _subfolderSelected[i] = sel;
                    filtersChanged();
                }
            }
            EditorGUILayout.EndScrollView();
            EditorGUILayout.EndVertical();
        }

        EditorGUILayout.Space(2);

        bool prevChoose = _chooseConcreteScripts;
        _chooseConcreteScripts = EditorGUILayout.ToggleLeft("Choose concrete scripts", _chooseConcreteScripts);
        if (_chooseConcreteScripts != prevChoose) filtersChanged();

        if (_chooseConcreteScripts)
        {
            if (_scriptFiles.Length == 0)
            {
                EditorGUILayout.HelpBox("No scripts found with current filters.", MessageType.Info);
            }
            else
            {
                EditorGUILayout.LabelField("Select scripts", EditorStyles.boldLabel);
                EditorGUILayout.BeginVertical(EditorStyles.helpBox);
                _scriptsScroll = EditorGUILayout.BeginScrollView(_scriptsScroll, GUILayout.Height(150));
                for (int i = 0; i < _scriptFiles.Length; i++)
                {
                    bool sel = EditorGUILayout.ToggleLeft(_scriptFiles[i], _scriptSelected[i]);
                    if (sel != _scriptSelected[i])
                    {
                        _scriptSelected[i] = sel;
                        invalidate();                
                    }
                }
                EditorGUILayout.EndScrollView();
                EditorGUILayout.EndVertical();
            }
        }
    }

    private void DrawOutputSelector(Action invalidate)
    {
        EditorGUILayout.LabelField("Output folder", EditorStyles.boldLabel);
        EditorGUILayout.BeginHorizontal();
        string newOut = EditorGUILayout.TextField(_outputFolder);
        if (newOut != _outputFolder) { _outputFolder = NormalizePath(newOut); invalidate(); }
        if (GUILayout.Button("…", GUILayout.Width(28)))
        {
            string sel = EditorUtility.OpenFolderPanel("Select output folder", _outputFolder, "");
            if (!string.IsNullOrEmpty(sel)) { _outputFolder = NormalizePath(sel); invalidate(); }
        }
        EditorGUILayout.EndHorizontal();

        EditorGUILayout.BeginHorizontal();
        GUILayout.Label("File name", GUILayout.Width(70));
        string newFile = EditorGUILayout.TextField(_outputFileName);
        if (newFile != _outputFileName)
        {
            _outputFileName = newFile.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                              ? newFile
                              : newFile + ".docx";
            invalidate();
        }
        EditorGUILayout.EndHorizontal();
    }

    private void DrawSuccessBox()
    {
        if (!_generationSucceeded) return;

        var ok = new GUIStyle(EditorStyles.label)
        { normal = { textColor = Color.green }, fontStyle = FontStyle.Bold };

        EditorGUILayout.BeginHorizontal(EditorStyles.helpBox);
        GUILayout.Label("✔ Word document saved", ok);
        GUILayout.FlexibleSpace();
        if (GUILayout.Button("Open Folder", GUILayout.Width(100)))
            EditorUtility.RevealInFinder(_generatedPath);
        EditorGUILayout.EndHorizontal();
    }

    private void RefreshSubfolders()
    {
        if (!Directory.Exists(_rootFolder))
        {
            _subfolders        = Array.Empty<string>();
            _subfolderSelected = Array.Empty<bool>();
            return;
        }

        _subfolders = Directory.GetDirectories(_rootFolder, "*", SearchOption.TopDirectoryOnly)
                               .Select(Path.GetFileName)
                               .ToArray();
        if (_subfolderSelected.Length != _subfolders.Length)
            _subfolderSelected = _subfolders.Select(_ => true).ToArray();
    }

    private void RefreshScripts()
    {
        if (!_chooseConcreteScripts || !Directory.Exists(_rootFolder))
        {
            _scriptFiles    = Array.Empty<string>();
            _scriptSelected = Array.Empty<bool>();
            return;
        }

        IEnumerable<string> files = Directory.GetFiles(_rootFolder, "*.cs", SearchOption.AllDirectories);

        if (!_includeAllSubfolders)
        {
            var allowed = new HashSet<string>(
                _subfolders.Where((s, i) => _subfolderSelected[i])
            );

            files = files.Where(p =>
            {
                var rel = p.Substring(_rootFolder.Length)
                           .TrimStart(Path.DirectorySeparatorChar, '/')
                           .Replace("\\", "/");
                var first = rel.Split('/')[0];
                return allowed.Contains(first);
            });
        }

        if (_excludeEditorFolders)
            files = files.Where(f => !f.Replace("\\", "/").Contains("/Editor/"));

        string[] newList = files
            .Select(p => p.Substring(_rootFolder.Length + 1).Replace("\\", "/"))
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (newList.Length == _scriptFiles.Length)
        {
            var dict = new Dictionary<string, bool>(_scriptFiles.Length, StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < _scriptFiles.Length; i++)
                dict[_scriptFiles[i]] = _scriptSelected[i];

            _scriptSelected = newList.Select(f => dict.TryGetValue(f, out bool v) ? v : true).ToArray();
        }
        else
        {
            _scriptSelected = newList.Select(_ => true).ToArray();
        }

        _scriptFiles = newList;
    }

    private void GenerateDocx()
    {
        _generationSucceeded = false;

        string[] files = Directory.GetFiles(_rootFolder, "*.cs", SearchOption.AllDirectories);

        if (!_includeAllSubfolders)
        {
            var allowed = new HashSet<string>(
                _subfolders.Where((s, i) => _subfolderSelected[i])
            );
            files = files.Where(p =>
            {
                var rel = p.Substring(_rootFolder.Length)
                           .TrimStart(Path.DirectorySeparatorChar, '/')
                           .Replace("\\", "/");
                var first = rel.Split('/')[0];
                return allowed.Contains(first);
            }).ToArray();
        }

        if (_excludeEditorFolders)
            files = files.Where(f => !f.Replace("\\", "/").Contains("/Editor/")).ToArray();

        if (_chooseConcreteScripts)
        {
            var chosen = new HashSet<string>(
                _scriptFiles.Where((rel, i) => _scriptSelected[i])
                            .Select(rel => NormalizePath(Path.Combine(_rootFolder, rel))),
                StringComparer.OrdinalIgnoreCase);

            files = files.Where(chosen.Contains).ToArray();
        }

        if (files.Length == 0)
        {
            EditorUtility.DisplayDialog("No scripts", "No C# scripts matched your selection.", "OK");
            return;
        }

        string docxPath = NormalizePath(Path.Combine(_outputFolder, _outputFileName));

        if (File.Exists(docxPath))
        {
            try { File.Delete(docxPath); }
            catch (IOException)
            {
                if (!EditorUtility.DisplayDialog("File in use",
                        "Close the file in Word and press Retry.", "Retry", "Cancel"))
                    return;
                File.Delete(docxPath);
            }
        }

        using var fs  = new FileStream(docxPath, FileMode.Create, FileAccess.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);

        AddEntry(zip, "[Content_Types].xml", GetContentTypesXml());
        AddEntry(zip, "_rels/.rels",        GetRootRelsXml());

        var scriptData = files.Select(p => (Path.GetFileName(p), File.ReadAllText(p)));
        AddEntry(zip, "word/document.xml", BuildDocumentXml(scriptData));

        zip.Dispose();
        AssetDatabase.Refresh();

        _generationSucceeded = true;
        _generatedPath       = docxPath;
        Debug.Log($"Collected {files.Length} scripts → '{docxPath}'");
    }

    /* ─────────── Служебные методы ─────────── */

    private static string NormalizePath(string p)
    {
        if (string.IsNullOrWhiteSpace(p)) return string.Empty;
        string full = Path.GetFullPath(p).Replace('/', '\\');
        return full.TrimEnd('\\');
    }

    private static void AddEntry(ZipArchive zip, string path, string content)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static string GetContentTypesXml() => @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml""  ContentType=""application/xml""/>
  <Override PartName=""/word/document.xml""
            ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
</Types>";

    private static string GetRootRelsXml() => @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""R1""
                Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""
                Target=""word/document.xml""/>
</Relationships>";

    private static string BuildDocumentXml(IEnumerable<(string name, string code)> scripts)
    {
        var sb = new StringBuilder();
        sb.Append(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>");

        foreach (var (name, code) in scripts)
        {
            AppendParagraph(sb, name);
            foreach (var line in code.Split('\n'))
                AppendParagraph(sb, line.Replace("\r", string.Empty));
        }

        sb.Append(@"
    <w:sectPr/>
  </w:body>
</w:document>");
        return sb.ToString();
    }

    private static void AppendParagraph(StringBuilder sb, string text)
    {
        string safe = SecurityElement.Escape(text);
        sb.Append($@"<w:p><w:r><w:t xml:space=""preserve"">{safe}</w:t></w:r></w:p>");
    }
}
#endif
