using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using Excel;
using System.Data;
using Newtonsoft.Json.Linq;
using System;

public class ExcelToJsonEditor 
{
    [MenuItem("Tools/Excel2Json Package2AB")]
    public static void CreateWindows()
    {
        EditorWindow.GetWindow(typeof(ExcelToJson));
    }
}


public class ExcelToJson : EditorWindow
{
    //Excel路径
    static string ExcelPath = "";
    //生成的json文件路径
    static string JsonPath = "";
    //打成AB包的路径
    static string ToABPath = "";

    /// <summary>
    /// 模板存放位置
    /// </summary>
    static string scriptsPath = "/Config/";


    /// <summary>
    /// 表格数据列表
    /// </summary>
    static List<TableData> dataList = new List<TableData>();

    void OnGUI()
    {
        selectPath("Excel文件路径", ref ExcelPath);
        selectPath("生成的json文件路径", ref JsonPath);
        selectPath("打成AB包的路径", ref ToABPath);
        Excel2Json(ExcelPath, JsonPath);


    }

    /// <summary>
    /// 选择各种路径
    /// </summary>
    private static void selectPath(string LabelName,ref string str)
    {
        GUILayout.Label(LabelName+":", EditorStyles.boldLabel);
        EditorGUILayout.BeginHorizontal(); // 开始一个水平布局组，确保路径文本框和按钮在同一行
        GUILayout.Label(str);  // 显示当前选择的路径
        if (GUILayout.Button("选择"+ LabelName)) // 绘制一个按钮，点击后打开文件夹选择对话框
        {
            // 打开文件夹选择对话框，返回所选路径并赋值给ExcelPath
            str = EditorUtility.OpenFolderPanel("Select Folder", "", "");
        }
        EditorGUILayout.EndHorizontal(); // 结束水平布局组

        SpaceSkin();
    }


    private static void Excel2Json(string ExcelPath,string JsonPath)
    {
        
        if (ExcelPath.Equals("")|| JsonPath.Equals(""))
        {
            GUILayout.Label("必要的文件路径为空！！！！", EditorStyles.boldLabel);
            GUILayout.Space(15);
        }

        if (GUILayout.Button("Excel转Json"))
        {
            ReadExcel(ExcelPath, JsonPath);
            AssetDatabase.Refresh();
        }

    }

    /// <summary>
    /// 分隔符
    /// </summary>
    private static void SpaceSkin()
    {
        GUILayout.Space(10);
        GUILayout.Label("", GUI.skin.horizontalSlider);
        GUILayout.Space(15);
    }


    public static void ReadExcel(string ExcelPath, string JsonPath)
    {
        if (Directory.Exists(ExcelPath))
        {
            //获取指定目录下所有的文件
            DirectoryInfo direction = new DirectoryInfo(ExcelPath);
            FileInfo[] files = direction.GetFiles("*.xlsx", SearchOption.AllDirectories);
            Debug.Log("文件数量:" + files.Length);

            for (int i = 0; i < files.Length; i++)
            {
                //if (files[i].Name.EndsWith(".meta") || !files[i].Name.EndsWith(".xlsx"))
                //{
                //    continue;
                //}
                Debug.Log("文件名称:" + files[i].FullName);
                LoadData(files[i].FullName, files[i].Name);
            }
        }
        else
        {
            Debug.LogError("当前选择的路径不存在Excel文件!");
        }
    }



    /// <summary>
    /// 读取表格并保存脚本及json
    /// </summary>
    static void LoadData(string filePath, string fileName)
    {
        //获取文件流
        FileStream fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        //生成表格的读取
        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
        // 表格数据全部读取到result里(引入：DataSet（using System.Data;）
        DataSet result = excelDataReader.AsDataSet();

        CreateTemplate(result, fileName);

        CreateJson(result, fileName);
    }



    /// <summary>
    /// 生成实体类模板
    /// </summary>
    static void CreateTemplate(DataSet result, string fileName)
    {
        if (!Directory.Exists(UnityEngine.Application.dataPath + scriptsPath))
        {
            Directory.CreateDirectory(UnityEngine.Application.dataPath + scriptsPath);
        }

        field = "";
        for (int i = 0; i < result.Tables[0].Columns.Count; i++)
        {
            string typeStr = result.Tables[0].Rows[1][i].ToString();
            typeStr = typeStr.ToLower();
            if (typeStr.Contains("[]"))
            {
                typeStr = typeStr.Replace("[]", "");
                typeStr = string.Format(" List<{0}>", typeStr);
            }


            //支持一些Unity的类型,例如Vector2-4
            //if (typeStr.Contains("vector"))
            //{
            //    typeStr = typeStr.Replace("v", "V");
            //}

            string nameStr = result.Tables[0].Rows[2][i].ToString();
            if (string.IsNullOrEmpty(typeStr) || string.IsNullOrEmpty(nameStr)) continue;
            field += "public " + typeStr + " " + nameStr + " { get; set; }\r\t\t";
        }
        fileName = fileName.Replace(".xlsx", "");
        Debug.Log(Eg_str);
        string tempStr = Eg_str;
        tempStr = tempStr.Replace("@Name", fileName);
        tempStr = tempStr.Replace("@File1", field);
        tempStr = tempStr.Replace("@JsonLastPath", (JsonPath + "/").ToString());
        tempStr = tempStr.Replace("@type", result.Tables[0].Rows[1][0].ToString());
        tempStr= tempStr.Replace("@variable", result.Tables[0].Rows[2][0].ToString());
        File.WriteAllText(UnityEngine.Application.dataPath + scriptsPath + fileName + ".cs", tempStr);

    }


    /// <summary>
    /// 生成json文件
    /// </summary>
    static void CreateJson(DataSet result, string fileName)
    {
        // 获取表格有多少列 
        int columns = result.Tables[0].Columns.Count;
        // 获取表格有多少行 
        int rows = result.Tables[0].Rows.Count;

        TableData tempData;
        string value;
        JArray array = new JArray();

        //第一行为表头，第二行为类型 ，第三行为字段名 不读取
        for (int i = 3; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                // 获取表格中指定行指定列的数据 
                value = result.Tables[0].Rows[i][j].ToString();

                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }
                tempData = new TableData();
                tempData.type = result.Tables[0].Rows[1][j].ToString();
                tempData.fieldName = result.Tables[0].Rows[2][j].ToString();
                tempData.value = value;

                dataList.Add(tempData);
            }

            if (dataList != null && dataList.Count > 0)
            {
                JObject tempo = new JObject();
                foreach (var item in dataList)
                {
                    switch (item.type)
                    {
                        case "string":
                            tempo[item.fieldName] = GetValue<string>(item.value);
                            break;
                        case "int":
                            tempo[item.fieldName] = GetValue<int>(item.value);
                            break;
                        case "float":
                            tempo[item.fieldName] = GetValue<float>(item.value);
                            break;
                        case "bool":
                            tempo[item.fieldName] = GetValue<bool>(item.value);
                            break;
                        case "string[]":
                            tempo[item.fieldName] = new JArray(GetList<string>(item.value, ','));
                            break;
                        case "int[]":
                            tempo[item.fieldName] = new JArray(GetList<int>(item.value, ','));
                            break;
                        case "float[]":
                            tempo[item.fieldName] = new JArray(GetList<float>(item.value, ','));
                            break;
                        case "bool[]":
                            tempo[item.fieldName] = new JArray(GetList<bool>(item.value, ','));
                            break;
                        //case "Vector2":
                        //    tempo[item.fieldName] = new JArray(GetList<float>(item.value, ','));
                        //    break;
                        //case "Vector3":
                        //    tempo[item.fieldName] = new JArray(GetList<float>(item.value, ','));
                        //    break;
                        //case "Vector4":
                        //    tempo[item.fieldName] = new JArray(GetList<float>(item.value, ','));
                        //    break;
                        default:
                            Debug.LogError($"该类型{item.type}暂不支持!");
                            break;
                    }
                }

                if (tempo != null)
                    array.Add(tempo);
                dataList.Clear();
            }
        }

        JObject o = new JObject();
        o["datas"] = array;
        o["version"] = "20200331";
        fileName = fileName.Replace(".xlsx", "");
        Debug.Log(JsonPath);
        File.WriteAllText(JsonPath+"/" + fileName + ".json", o.ToString());
    }

  

    /// <summary>
    /// 字段
    /// </summary>
    static string field;

    /// <summary>
    /// 实体类模板
    /// </summary>
    static string Eg_str =

        "using System.Collections.Generic;\r" +
        "using UnityEngine;\r" +
        "using System.IO;\r" +
        "using Newtonsoft.Json;\r\r" +
        "public class @Name  {\r\r\t\t" +
        "@File1 \r\t\t" +
        "public static string configName = \"@Name\";\r\t\t" +
        "public static @Name config { get; set; }\r\t\t" +
        "public string version { get; set; }\r\t\t" +
        "public List<@Name> datas { get; set; }\r\r\t\t" +
        "public static void Init()\r\t\t{\r\t\t\tstring folderPath = \"@JsonLastPath\";\r\t\t\tstring[] filePaths = Directory.GetFiles(folderPath, configName + \".json\");\r\t\t\tif (filePaths != null)\r\n\t\t\t{\r\n\t\t\t\tstring jsonContent = File.ReadAllText(filePaths[0]);\r\n\t\t\t\tconfig = JsonConvert.DeserializeObject<@Name>(jsonContent);\r\t\t\t }\r\r\t\t}\r\r\t\t" +
        "public static @Name Get(@type @variable)\r\t\t{\r\t\t\tInit();\r\t\t\tforeach (var item in config.datas)\r\t\t\t{\r\t\t\t\tif (item.@variable == @variable)\r\t\t\t\t{ \r\t\t\t\t\treturn item;\r\t\t\t\t}\r\t\t\t}\r\t\t\treturn null;\r\t\t}\r"
         + "\r}";



    /// <summary>
    /// 字符串拆分列表
    /// </summary>
    static List<T> GetList<T>(string str, char spliteChar)
    {
        string[] ss = str.Split(spliteChar);
        int length = ss.Length;
        List<T> arry = new List<T>(ss.Length);
        for (int i = 0; i < length; i++)
        {
            arry.Add(GetValue<T>(ss[i]));
        }
        return arry;
    }

    static T GetValue<T>(object value)
    {
        return (T)Convert.ChangeType(value, typeof(T));
    }


    /// <summary>
    /// 特殊类型 例如Vector2-4
    /// </summary>
    //static IFormattable GetVectorValue(string str,string spliteChar)
    //{
    //    string[] ss = str.Split(spliteChar);
    //    int length = ss.Length;
    //    switch (length)
    //    {
    //        case 2:
    //            Vector2 vector = new Vector2();
    //            float.TryParse(ss[0], out vector.x);
    //            float.TryParse(ss[1], out vector.y);
    //            return vector;
    //        case 3:
    //            Vector3 vector1 = new Vector3();
    //            float.TryParse(ss[0], out vector1.x);
    //            float.TryParse(ss[1], out vector1.y);
    //            float.TryParse(ss[2], out vector1.z);
    //            return vector1;
    //        case 4:
    //            Vector4 vector2 = new Vector4();
    //            float.TryParse(ss[0], out vector2.x);
    //            float.TryParse(ss[1], out vector2.y);
    //            float.TryParse(ss[2], out vector2.z);
    //            float.TryParse(ss[3], out vector2.w);
    //            return vector2;
    //    }
    //    return null;
    //}
       


    public struct TableData
    {
        public string fieldName;
        public string type;
        public string value;

        public override string ToString()
        {
            return string.Format("fieldName:{0} type:{1} value:{2}", fieldName, type, value);
        }
    }


}