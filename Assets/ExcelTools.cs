using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using OfficeOpenXml;

public class ExcelTools : MonoBehaviour {

	// Use this for initialization
	void Start () {
		
	}
	
	// Update is called once per frame
	void Update () {
		
	}

    [MenuItem("1/1")]
    static void CreateExcelTable()
    {
        Debug.Log("CreateExcelTable test");
        string path = Application.dataPath + "test.xlsx";
        FileStream fs = new FileStream(path, FileMode.Create);

        using (var package = new ExcelPackage(fs))
        {
            Debug.Log("CreateExcelTable test 22");
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("sheet1");
            worksheet.Cells[1, 1].Value = "Company Name1";
            worksheet.Cells[1, 2].Value = "Address1";

            worksheet = package.Workbook.Worksheets.Add("sheet2");
            worksheet.Cells[1, 1].Value = "Company Name2";
            worksheet.Cells[1, 2].Value = "Address2";
            package.Save();
        }
        fs.Close();
        AssetDatabase.Refresh();
    }
}
