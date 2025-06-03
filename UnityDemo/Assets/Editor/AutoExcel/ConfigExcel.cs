using UnityEngine;
using System.Collections;
using UnityEditor;
using System.IO;
using System.Collections.Generic;
using System;
using System.Linq;

public class ConfigExcel : EditorWindow
{
    [MenuItem("Tools/ConfigExcel")]
    public static void ShowWindow()
    {
        var thisWindow = GetWindow(typeof(ConfigExcel));
        thisWindow.titleContent = new GUIContent("ConfigExcel");
    }
    string inputDir
    {
        get => PlayerPrefs.GetString("ConfigExcel_inputDir");
        set => PlayerPrefs.SetString("ConfigExcel_inputDir", value);
    }
    string outputDir
    {
        get => PlayerPrefs.GetString("ConfigExcel_outputDir");
        set => PlayerPrefs.SetString("ConfigExcel_outputDir", value);
    }
    string ignoreFiles
    {
        get => PlayerPrefs.GetString("ConfigExcel_ignoreFiles", "ignore1.xlsx,ignore2.xlsx");
        set => PlayerPrefs.SetString("ConfigExcel_ignoreFiles", value);
    }
    void OnGUI()
    {
        inputDir = EditorGUILayout.TextField("输入目录", inputDir);
        outputDir = EditorGUILayout.TextField("输出目录", outputDir);
        ignoreFiles = EditorGUILayout.TextField("忽略文件(英文逗号隔开)", ignoreFiles);
        if (GUILayout.Button("导出Excel", GUILayout.Height(30)))
        {
            Excel2Code.Excel2Code.GenerateFromDir(inputDir, ignoreFiles.Trim().Split(',', StringSplitOptions.RemoveEmptyEntries).ToList(), outputDir);
        }
    }
}
