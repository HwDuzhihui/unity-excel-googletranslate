#if UNITY_EDITOR
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.Networking;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using UnityEditor;

public class GoogleTranslate : MonoBehaviour
{
    static GoogleTranslate _Instance;
    static GoogleTranslate Instance
    {
        get
        {
            if (_Instance == null)
            {
                _Instance = new GameObject("GoogleTranslate").AddComponent<GoogleTranslate>();
            }
            return _Instance;
        }
    }
    //语言简写--对应语言表的语言列
    static List<string> language = new List<string>()
    {
        "zh-cn", //简中
        "zh-tw", //繁中
        "en",    //英语
        "ja",    //日语
        "ko",    //韩语
        "de",    //德语
        "fr",    //法语
        "es",    //西班牙
        "it",    //意大利语
        "pt",    //葡萄牙语
        "tr",    //土耳其语
        "id",    //印尼语
        "th",    //泰语
        "vi",    //越南语
        "ar",    //阿拉伯语
        "ru",    //俄语
    };
    //翻译不支持一些特殊字符,需要翻译前转一下,翻译后转回来
    static Dictionary<string, string> ignore = new Dictionary<string, string>()
    {
        {"#",           "_GTCHAR1_" },
        {"\\n",         "_GTCHAR2_" },
        {"\n",          "_GTCHAR3_" },
        {"+",           "_GTCHAR4_" },
        {"%",           "_GTCHAR5_" },
    };
    //语言表目录
    static string excelFile = Application.dataPath.Substring(0, Application.dataPath.Length - 7) + "/Excel/语言表.xlsx";
    //多语言列
    static List<List<string>> sheetData = new List<List<string>>();
    //需要翻译的列
    static Dictionary<int,Dictionary<int,string>> translateIndex = new Dictionary<int, Dictionary<int, string>>();
    //需要翻译多少次
    static int translateMax = 0;
    //当前翻译次数
    static int translateCur = 0;

    [MenuItem("工具栏/翻译语言表")]
    public static void TranslateExcel()
    {
        if(!Application.isPlaying)
        {
            Debug.LogError("在游戏运行时翻译");
            return;
        }
        if (!File.Exists(excelFile))
        {
            Debug.LogError("不存在:" + excelFile);
            return;
        }
        sheetData.Clear();
        translateIndex.Clear();
        translateMax = 0;
        translateCur = 0;
        using (FileStream fileStream = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(0);
            int rowMax = sheet.LastRowNum + 1;
            int columnMax = sheet.GetRow(0).LastCellNum;
            for (int m = 0; m < columnMax; m++)
            {
                ICell cell = sheet.GetRow(0).GetCell(m);
                List<string> cc = new List<string>();
                for (int n = 0; n < rowMax; n++)
                {
                    IRow row = sheet.GetRow(n);
                    if (row == null) break;
                    cell = row.GetCell(m);
                    string msg = cell == null ? "" : cell.ToString();
                    cc.Add(msg);
                }
                sheetData.Add(cc);
            }
            //需要翻译的列
            for (int i = 2; i < sheetData.Count; i++)
            {
                for (int j = 0; j < sheetData[i].Count; j++)
                {
                    if (string.IsNullOrEmpty(sheetData[i][j]))
                    {
                        if(!translateIndex.ContainsKey(i))
                        {
                            translateIndex.Add(i, new Dictionary<int, string>());
                        }
                        translateIndex[i].Add(j, "");
                        translateMax++;
                    }
                }
            }
            if(translateMax > 0)
            {
                Debug.Log("开始翻译");
                foreach (var kv1 in translateIndex)
                {
                    foreach (var kv2 in kv1.Value)
                    {
                        //需要翻译哪种语言
                        int i = kv1.Key;
                        //需要翻译哪一行
                        int j = kv2.Key;
                        Instance.StartCoroutine(TranslateAsync(i, j));
                    }
                }
            }
            else
            {
                Debug.Log("无需翻译");
            }
        }
    }
    static IEnumerator TranslateAsync(int i,int j)
    {
        //要翻译的语言
        string lg = language[i - 1];
        //需要翻译的中文
        string cn = sheetData[1][j];
        //转义无法翻译的字符
        foreach (var kv in ignore)
        {
            if (cn.Contains(kv.Key))
            {
                cn = cn.Replace(kv.Key, kv.Value);
            }
        }
        string requestUrl = string.Format("https://translate.googleapis.com/translate_a/single?client=gtx&sl={0}&tl={1}&dt=t&q={2}", "zh-cn", lg, cn);
        UnityWebRequest request = UnityWebRequest.Get(requestUrl);
        yield return request.SendWebRequest();
        if (string.IsNullOrEmpty(request.error))
        {
            JArray jsonArray = (JArray)JsonConvert.DeserializeObject(request.downloadHandler.text);
            JToken jToken = jsonArray[0];
            string translate = "";
            foreach(var kv in jToken)
            {
                translate += kv[0];
            }
            //转回来
            foreach (var kv in ignore)
            {
                if (translate.Contains(kv.Value))
                {
                    translate = translate.Replace(kv.Value, kv.Key);
                }
            }
            Debug.Log(sheetData[1][j] + "  " + translate);
            translateIndex[i][j] = translate;
            sheetData[i][j] = translate;
        }
        else
        {
            Debug.LogError("翻译错误:" + request.error + "  " + language + " : " + cn);
        }
        translateCur++;
        if (translateCur == translateMax)
        {
            WriteToExcel();
        }
    }
    static void WriteToExcel()
    {
        IWorkbook workbook;
        using (FileStream fileStream = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(fileStream);
        }
        ISheet sheet = workbook.GetSheetAt(0);
        foreach (var kv1 in translateIndex)
        {
            int i = kv1.Key;
            foreach (var kv2 in kv1.Value)
            {
                int j = kv2.Key;
                IRow row = sheet.GetRow(j);
                ICell cell = row?.GetCell(i);
                if(row == null)
                {
                    row = sheet.CreateRow(j);
                }
                if( cell == null)
                {
                    cell = row.CreateCell(i);
                }
                cell?.SetCellValue(kv2.Value);
            }
        }
        using (FileStream fileStream = new FileStream(excelFile, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(fileStream);
        }
        Debug.Log("翻译完成");
    }
}
#endif
