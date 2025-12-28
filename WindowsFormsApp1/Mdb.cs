using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;//以chatGPT建立的
using System.IO;
using WindowsFormsApp1;
using System.Windows.Forms;
using System.Web.UI;
using OpenQA.Selenium.DevTools.V85.ApplicationCache;
using System.Windows.Media.Animation;
//引用adodb 要將其「內嵌 Interop 類型」（Embed Interop Type）屬性設為false（預設是true）才不會出現以下錯誤：  HResult=0x80131522  Message=無法從組件 載入類型 'ADODB.FieldsToInternalFieldsMarshaler'。
//https://stackoverflow.com/questions/5666265/adodbcould-not-load-type-adodb-fieldstointernalfieldsmarshaler-from-assembly  https://blog.csdn.net/m15188153014/article/details/119895082
using ado = ADODB;//https://docs.microsoft.com/zh-tw/dotnet/csharp/language-reference/keywords/using-directive

namespace TextForCtext
{
    class Mdb
    {
        static Form1 frm = Application.OpenForms["Form1"] as Form1;
        public static string DropBoxPathIncldBackSlash = getDropBoxPathIncldBackSlash();
        static string getDropBoxPathIncldBackSlash()
        {
            if (string.IsNullOrEmpty(DropBoxPathIncldBackSlash))
            {
                DropBoxPathIncldBackSlash = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Dropbox\";
                DropBoxPathIncldBackSlash = Directory.Exists(DropBoxPathIncldBackSlash) ? DropBoxPathIncldBackSlash : DropBoxPathIncldBackSlash.Replace(@"C:\", @"A:\");
            }
            return DropBoxPathIncldBackSlash;
        }
        internal static string fileFullName(string dbNameIncludeExt)
        {
            if (frm == null) frm = Application.OpenForms["Form1"] as Form1;
            string root = frm.dropBoxPathIncldBackSlash;
            if (!File.Exists(root + dbNameIncludeExt))
            {
                root = root.Replace("C:", "A:");
            }
            frm = null;
            if (!File.Exists(root + dbNameIncludeExt)) { MessageBox.Show(root + dbNameIncludeExt + "not found"); return ""; }
            else
                return root + dbNameIncludeExt;
        }

        /// <summary>
        /// ado.Connection
        /// </summary>
        /// <param name="dbNameIncludeExt"></param>
        /// <param name="cnt"></param>
        internal static void openDatabase(string dbNameIncludeExt, ref ado.Connection cnt)
        {
            string root = getDropBoxPathIncldBackSlash();//DropBoxPathIncldBackSlash;
            if (!File.Exists(root + dbNameIncludeExt))
            {
                root = root.Replace("C:", "A:");
            }
            if (!File.Exists(root + dbNameIncludeExt)) { MessageBox.Show(root + dbNameIncludeExt + "not found"); return; }
            //string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + root + dbNameIncludeExt;
            string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + root + dbNameIncludeExt;
            try
            {

                cnt.Open(conStr);
            }
            catch (Exception)
            {

                try
                {
                    //conStr = conStr.Replace("Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0");
                    conStr = conStr.Replace("Microsoft.Jet.OLEDB.4.0", "Microsoft.ACE.OLEDB.12.0");
                    cnt.Open(conStr);

                }
                catch (Exception)
                {

                    throw;
                }


            }
        }
        /// <summary>
        /// OleDbConnection
        /// </summary>
        /// <param name="dbFileFullname"></param>
        /// <param name="conn"></param>
        private static void openDb(string dbFileFullname, out OleDbConnection conn)
        {
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbFileFullname;

            // 建立連接物件            
            conn = new OleDbConnection(connectionString);

            // 開啟資料庫連接
            conn.Open();
        }

        /// <summary>
        /// 檢查異體字轉正之資料存在否
        /// </summary>
        /// <param name="wordtoChk">要檢查的異體字</param>
        /// <returns></returns>
        internal static bool VariantsExist(string wordtoChk)//以chatGPT建立再自己略加修潤的 Alt + v:即以以下與chatGPT對話所得者：C# 檢查[查字.mdb].[異體字反正]資料表中是否已有該字記錄,擬自創 creedit 一動詞以作紀念，日後若有標識 creedit（creeditted 、 creeditting) 者，即為取自 chatGPT AI 而改寫者，意為：「create from chatGPT AI and edit」,以取 create 諧音且兼其義以識別非純自創也 感恩感恩　讚歎讚歎　南無阿彌陀佛 
        {//20221231,心得感觸啟發可略見此（末學臉書）：https://www.facebook.com/oscarsun72/posts/pfbid0TXr2QwArfHcL3XqsFHMg8cFbzj8zd2fBzzoMXermrrNqXccb626hfZasb6hB1p7Ql
            /*
                我反而更想寫了，因為AI可以義務地作我的小編，無怨無悔又不會曠職摸魚偷懶地幫我先完成前置作業，我這個主編或編審再決定要不要採用小編的建議與方案。畢竟最後拍板權還在自己、定奪璽還在自己的肉身手上，且不會收到小編們任何的怨懟與不滿，何樂而不為呢？友直友諒友多聞，學而時習之，不亦說乎；有朋自遠方來，不亦樂乎；人不知而不慍，不亦君子乎。感恩感恩　讚歎讚歎　南無阿彌陀佛
                剛才得力於chatGPT後的心得
                人往高處爬，本來基本呆板吃力費神的繁瑣就該讓嘍囉作，可誰肯甘心情願作我的嘍囉小的呢？能得此不會抱怨、沒有情緒、不會請假生病、藉口假裝的AI助理，想是任何想要身在最高層的人類都夢寐以求的無償坐擁吧。把全副的智慧與精神耗費在真正AI無法完成的工作、志業，不也是咱們人類本當如是的生活與萬物之靈的生命意義麼。AI的出現，正好淘汰吾人耍廢偷懶自暴自棄的性格，讓咱們向上精進的鬥志由此而被激發，不是諍友諫友良師益友，而是什麼呢？
                本來，高手不在他用的工具較一般人優利，而是他能善用一切有用可用的工具，如一葦即能渡江，木劍也勝刀劍，化腐朽為神奇，太極而無極，無來亦無去……。感恩感恩　南無阿彌陀佛
                https://ctext.org/library.pl?if=gb&file=78287&page=128#%E8%BA%AB%E5%9C%A8%E6%9C%80%E9%AB%98%E5%B1%A4
                為什麼不這樣想呢？：功夫再好，再加上會用子彈呢？如虎添翼，站在巨人的肩膀上（=聖人、賢人或能人，如劉邦、劉備，手下誰不比他強，但他佬，主編最強），是不是這麼說的呢？轉個念頭，世界不同，心能轉物，則同如來。我們不用抵抗，而是要內化為我的功夫之一，否則就像滿清末年（一國末年如一身老年），維新變法，最後只能亡國絕祀了。感恩感恩 南無阿彌陀佛
            */
            // 建立連接字串
            string f = fileFullName("查字.mdb");
            if (f == "")
            {
                MessageBox.Show("找不到「查字.mdb」");
                return false;
            }
            openDb(f, out OleDbConnection conn);
            // 建立命令物件
            OleDbCommand cmd = conn.CreateCommand();

            // 設定命令的文字
            //cmd.CommandText = "SELECT COUNT(*) FROM 異體字轉正 WHERE 異體字 = @word";//@word 為參數名，用「=」比對中文會不正確，在cjk-擴充字集時
            cmd.CommandText = "SELECT COUNT(*) FROM 異體字轉正 WHERE strcomp(異體字 , @word)=0";//@word 為參數名

            // 設定命令的參數
            cmd.Parameters.AddWithValue("@word", wordtoChk);

            // 執行命令並取得結果
            int count = (int)cmd.ExecuteScalar();
            //*/

            // 關閉資料庫連接
            conn.Close();
            // 判斷結果
            if (count > 0)
            {
                return true;// 資料表中已有該字記錄
            }
            else
            {
                Clipboard.SetText(wordtoChk);//複製到剪貼簿以便到MS Access 輸入新記錄時直接貼上
                return false;// 資料表中沒有該字記錄
            }

        }

        /// <summary>
        /// 檢查目前IP的狀態
        /// 由 VariantsExist 改寫而來 20231224平安夜 由Bing大菩薩所改寫者
        /// </summary>
        /// <param name="iptoChk">要檢查的IP</param>
        /// <returns>回傳一個Tuple分別對應IP資料表除了IP欄位外的各個欄位值
        /// IpAddressBanned、IPisblocked、ctext、RecordDate
        /// 若沒找到IP，則傳回null
        /// </returns>
        internal static Tuple<bool, bool, bool, bool, DateTime> IPStatus(string iptoChk)
        {
            // 建立連接字串
            string f = fileFullName("查字.mdb");
            if (f == "")
            {
                MessageBox.Show("找不到「查字.mdb」");
                return null;
            }
            openDb(f, out OleDbConnection conn);
            // 建立命令物件
            OleDbCommand cmd = conn.CreateCommand();

            // 設定命令的文字
            cmd.CommandText = "SELECT * FROM IP WHERE strcomp(IP , @IP)=0";//@word 為參數名

            // 設定命令的參數
            cmd.Parameters.AddWithValue("@IP", iptoChk);

            // 執行命令並取得結果
            OleDbDataReader reader = cmd.ExecuteReader();

            // 判斷結果
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    bool IpAddressBanned = reader.GetBoolean(reader.GetOrdinal("IpAddressBanned"));
                    bool IPisblocked = reader.GetBoolean(reader.GetOrdinal("IPisblocked"));
                    bool ctext = reader.GetBoolean(reader.GetOrdinal("ctext"));
                    bool Systemisbusy = reader.GetBoolean(reader.GetOrdinal("Systemisbusy"));
                    DateTime RecordDate = reader.GetDateTime(reader.GetOrdinal("RecordDate"));
                    reader.Close();
                    conn.Close();
                    return new Tuple<bool, bool, bool, bool, DateTime>(IpAddressBanned, IPisblocked, ctext, Systemisbusy, RecordDate);
                }
            }
            else
            {
                try
                {
                    Clipboard.SetText(iptoChk);//複製到剪貼簿以便到MS Access 輸入新記錄時直接貼上
                }
                catch (Exception)
                {
                }
                reader.Close();
                conn.Close();
                return null;// 資料表中沒有該字記錄
            }
            reader.Close();
            conn.Close();
            return null;
        }

        //internal static Tuple<bool,bool,bool,DateTime>    IPStatusTemp(string iptoChk)
        //{
        //    // 建立連接字串
        //    string f = fileFullName("查字.mdb");
        //    if (f == "")
        //    {
        //        MessageBox.Show("找不到「查字.mdb」");
        //        return null;
        //    }
        //    openDb(f, out OleDbConnection conn);
        //    // 建立命令物件
        //    OleDbCommand cmd = conn.CreateCommand();

        //    // 設定命令的文字
        //    //cmd.CommandText = "SELECT COUNT(*) FROM 異體字轉正 WHERE 異體字 = @word";//@word 為參數名，用「=」比對中文會不正確，在cjk-擴充字集時
        //    cmd.CommandText = "SELECT COUNT(*) FROM IP WHERE strcomp(IP , @IP)=0";//@word 為參數名

        //    // 設定命令的參數
        //    cmd.Parameters.AddWithValue("@IP", iptoChk);

        //    // 執行命令並取得結果
        //    int count = (int)cmd.ExecuteScalar();
        //    //*/

        //    // 關閉資料庫連接
        //    conn.Close();
        //    // 判斷結果
        //    if (count > 0)
        //    {
        //        return new Tuple();// 資料表中已有該字記錄,則以Tuple回傳除了IP欄位外的各個欄位值
        //    }
        //    else
        //    {
        //        Clipboard.SetText(iptoChk);//複製到剪貼簿以便到MS Access 輸入新記錄時直接貼上
        //        return null;// 資料表中沒有該字記錄
        //    }

        //}





        /// <summary>
        /// 輸入平抬條件：0=後綴；1=前綴；2=前後之前；3前後之後；4是前+後之詞彙；5非前+後之詞彙；6非後綴之詞彙；7非前綴之詞彙
        /// Alt + l
        /// </summary>
        /// <param name="termtoChk"></param>
        internal static void TopLineFactorIuput04condition(string termtoChk)
        {
            //TextBox tb = (TextBox)frm.Controls["textBox1"];
            //string termtoChk = tb.SelectedText;

            //開啟"查字.mdb"資料庫
            string f = fileFullName("查字.mdb");
            if (f == "")
            {
                MessageBox.Show("找不到「查字.mdb」");
                return;
            }

            openDb(f, out OleDbConnection conn);

            //20230114 creedit chatGPT大菩薩：新增資料庫資料：
            //在這個程式碼中,使用了一個 T-SQL 的 IF @@ROWCOUNT = 0 判斷當前存在與否,如果不存在就新增,否則就查詢。
            //這樣子就可以同時達到查詢和新增的目的了。
            using (OleDbCommand cmd = conn.CreateCommand())
            {
                int condition = 0;//輸入平抬條件：0=後綴；1=前綴；2=前後之前；3前後之後；4是前+後之詞彙；5非前+後之詞彙；6非後綴之詞彙；7非前綴之詞彙
                if (termtoChk.IndexOf("|" + Environment.NewLine) > -1)
                {
                    termtoChk = termtoChk.Replace("|" + Environment.NewLine, "");
                    condition = 4;
                }
                //cmd.CommandText = "SELECT COUNT(*) FROM 每行字數判斷用 WHERE strcomp(term , @term)=0; " +
                //                  "IF @@ROWCOUNT = 0 " +
                //                  "INSERT INTO 每行字數判斷用 (term,condition) VALUES (@term, @condition);" +
                //                  "ELSE " +
                //"SELECT term,condition FROM 每行字數判斷用 WHERE strcomp(term , @term)=0;";
                cmd.CommandText = "SELECT term,condition FROM 每行字數判斷用 WHERE strcomp(term , @term)=0;";
                cmd.Parameters.AddWithValue("@term", termtoChk);
                //string input = Microsoft.VisualBasic.Interaction.InputBox("Prompt", "Title", "Default Value", -1, -1);                
                cmd.Parameters.AddWithValue("@condition", condition);
                string term_condition = "";
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    //if (reader.RecordsAffected == 0)
                    if (reader.HasRows)
                    {
                        //reader.NextResult();
                        while (reader.Read())
                        {
                            term_condition += reader["term"].ToString() + "=" +
                             reader["condition"].ToString() + Environment.NewLine;
                            // do something with term and condition
                        }
                    }
                    else
                    {
                        //if (DialogResult.OK == MessageBox.Show("確定新增？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation))
                        if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("確定新增「"+ termtoChk +"」？", "【抬頭資訊新增】"))
                        {
                            cmd.CommandText = "INSERT INTO 每行字數判斷用 (term,condition) VALUES (@term, @condition);";
                            reader.Close();
                            cmd.ExecuteReader().Close();
                        }
                    }
                }
                if (term_condition != "")
                    //MessageBox.Show("現有資料：" + Environment.NewLine + term_condition);
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("現有資料：" + Environment.NewLine + term_condition);
            }
            conn.Close();
            ////設定查詢指令
            //int count = 0;
            //using (OleDbCommand cmd = conn.CreateCommand())
            //{
            //    cmd.CommandText = "SELECT COUNT(*) FROM 每行字數判斷用 WHERE strcomp(term , @term)=0";//@word 為參數名
            //    cmd.Parameters.AddWithValue("@word", termtoChk);
            //    count = (int)cmd.ExecuteScalar();

            //    if (count == 0)
            //    {//新增 termtoChk 值 至 每行字數判斷用 資料表中
            //     //20230114 creedit chatGPT大菩薩：新增資料庫資料：
            //        using (OleDbCommand cmdAddNewRec = conn.CreateCommand())
            //        {
            //            cmdAddNewRec.CommandText = "INSERT INTO 每行字數判斷用 (term) VALUES (@term)";
            //            cmdAddNewRec.Parameters.AddWithValue("@term", termtoChk);
            //            cmdAddNewRec.ExecuteNonQuery();
            //        }

            //    }
            //    else
            //    {
            //        MessageBox.Show("term="+);
            //    }
            //}
        }
    }
}
