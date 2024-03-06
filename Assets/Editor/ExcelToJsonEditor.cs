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
    //Excel·��
    static string ExcelPath = "";
    //���ɵ�json�ļ�·��
    static string JsonPath = "";
    //���AB����·��
    static string ToABPath = "";

    /// <summary>
    /// ģ����λ��
    /// </summary>
    static string scriptsPath = "/Config/";


    /// <summary>
    /// ��������б�
    /// </summary>
    static List<TableData> dataList = new List<TableData>();

    void OnGUI()
    {
        selectPath("Excel�ļ�·��", ref ExcelPath);
        selectPath("���ɵ�json�ļ�·��", ref JsonPath);
        selectPath("���AB����·��", ref ToABPath);
        Excel2Json(ExcelPath, JsonPath);


    }

    /// <summary>
    /// ѡ�����·��
    /// </summary>
    private static void selectPath(string LabelName,ref string str)
    {
        GUILayout.Label(LabelName+":", EditorStyles.boldLabel);
        EditorGUILayout.BeginHorizontal(); // ��ʼһ��ˮƽ�����飬ȷ��·���ı���Ͱ�ť��ͬһ��
        GUILayout.Label(str);  // ��ʾ��ǰѡ���·��
        if (GUILayout.Button("ѡ��"+ LabelName)) // ����һ����ť���������ļ���ѡ��Ի���
        {
            // ���ļ���ѡ��Ի��򣬷�����ѡ·������ֵ��ExcelPath
            str = EditorUtility.OpenFolderPanel("Select Folder", "", "");
        }
        EditorGUILayout.EndHorizontal(); // ����ˮƽ������

        SpaceSkin();
    }


    private static void Excel2Json(string ExcelPath,string JsonPath)
    {
        
        if (ExcelPath.Equals("")|| JsonPath.Equals(""))
        {
            GUILayout.Label("��Ҫ���ļ�·��Ϊ�գ�������", EditorStyles.boldLabel);
            GUILayout.Space(15);
        }

        if (GUILayout.Button("ExcelתJson"))
        {
            ReadExcel(ExcelPath, JsonPath);
            AssetDatabase.Refresh();
        }

    }

    /// <summary>
    /// �ָ���
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
            //��ȡָ��Ŀ¼�����е��ļ�
            DirectoryInfo direction = new DirectoryInfo(ExcelPath);
            FileInfo[] files = direction.GetFiles("*.xlsx", SearchOption.AllDirectories);
            Debug.Log("�ļ�����:" + files.Length);

            for (int i = 0; i < files.Length; i++)
            {
                //if (files[i].Name.EndsWith(".meta") || !files[i].Name.EndsWith(".xlsx"))
                //{
                //    continue;
                //}
                Debug.Log("�ļ�����:" + files[i].FullName);
                LoadData(files[i].FullName, files[i].Name);
            }
        }
        else
        {
            Debug.LogError("��ǰѡ���·��������Excel�ļ�!");
        }
    }



    /// <summary>
    /// ��ȡ��񲢱���ű���json
    /// </summary>
    static void LoadData(string filePath, string fileName)
    {
        //��ȡ�ļ���
        FileStream fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        //���ɱ��Ķ�ȡ
        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
        // �������ȫ����ȡ��result��(���룺DataSet��using System.Data;��
        DataSet result = excelDataReader.AsDataSet();

        CreateTemplate(result, fileName);

        CreateJson(result, fileName);
    }



    /// <summary>
    /// ����ʵ����ģ��
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


            //֧��һЩUnity������,����Vector2-4
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
    /// ����json�ļ�
    /// </summary>
    static void CreateJson(DataSet result, string fileName)
    {
        // ��ȡ����ж����� 
        int columns = result.Tables[0].Columns.Count;
        // ��ȡ����ж����� 
        int rows = result.Tables[0].Rows.Count;

        TableData tempData;
        string value;
        JArray array = new JArray();

        //��һ��Ϊ��ͷ���ڶ���Ϊ���� ��������Ϊ�ֶ��� ����ȡ
        for (int i = 3; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                // ��ȡ�����ָ����ָ���е����� 
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
                            Debug.LogError($"������{item.type}�ݲ�֧��!");
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
    /// �ֶ�
    /// </summary>
    static string field;

    /// <summary>
    /// ʵ����ģ��
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
    /// �ַ�������б�
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
    /// �������� ����Vector2-4
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