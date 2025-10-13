using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Automatic_Storage
{
    class Log
    {
        public enum CRUD
        {
            Insert = 0,
            Select = 1,
            Update = 2,
            Delete = 3            
        }
        struct Action_Button
        {
            public string btn_Input;
            public string btn_Out;
            public string btn_BatIn;
            public string btn_BatOut;
            public string button1;                            //查詢全部
            public string btn_combi;                       //料號+儲位合併查詢
            public string btn_delPosition;             //儲位刪除
            public string btn_findAll;                      //歷史_全部查詢
            public string btn_itemSite;                  //歷史_料號+儲位合併查詢
            public string btn_reP2;                         //歷史_返回
            public string selectButton;                  //檔案選擇
            public string commitButton;              //檔案上傳
            public string btn_return;                      //檔案_介面返回

            public string maintainRecord;
            public string historyRecord;
        }
        struct Action_TextBox
        {
            public string itemRecord;
            public string positionRecord;
            
        }
        public void WriteLog(string EventName) 
        {
            
        }
    }
}
