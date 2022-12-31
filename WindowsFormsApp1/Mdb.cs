using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;//以chatGPT建立的

namespace TextForCtext
{
    class Mdb
    {
        internal bool VariantsExist()//以chatGPT建立再自己略加修潤的 Alt + v:即以以下與chatGPT對話所得者：C# 檢查[查字.mdb].[異體字反正]資料表中是否已有該字記錄,擬自創 creedit 一動詞以作紀念，日後若有標識 creedit（creeditted 、 creeditting) 者，即為取自 chatGPT AI 而改寫者，意為：「create from chatGPT AI and edit」,以取 create 諧音且兼其義以識別非純自創也 感恩感恩　讚歎讚歎　南無阿彌陀佛 
        {//20221231,心得感觸啟發可略見此（末學臉書）：https://www.facebook.com/oscarsun72/posts/pfbid0TXr2QwArfHcL3XqsFHMg8cFbzj8zd2fBzzoMXermrrNqXccb626hfZasb6hB1p7Ql
            /*
                我反而更想寫了，因為AI可以義務地作我的小編，無怨無悔又不會曠職摸魚偷懶地幫我先完成前置作業，我這個主編或編審再決定要不要採用小編的建議與方案。畢竟最後拍板權還在自己、定奪璽還在自己的肉身手上，且不會收到小編們任何的怨懟與不滿，何樂而不為呢？友直友諒友多聞，學而時習之，不亦說乎；有朋自遠方來，不亦樂乎；人不知而不慍，不亦君子乎。感恩感恩　讚歎讚歎　南無阿彌陀佛
                剛才得力於chatGPT後的心得
                人往高處爬，本來基本呆板吃力費神的繁瑣就該讓嘍囉作，可誰肯甘心情願作我的嘍囉小的呢？能得此不會抱怨、沒有情緒、不會請假生病、藉口假裝的AI助理，想是任何想要身在最高層的人類都夢寐以求的無償坐擁吧。把全副的智慧與精神耗費在真正AI無法完成的工作、志業，不也是咱們人類本當如是的生活與萬物之靈的生命意義麼。AI的出現，正好淘汰吾人耍廢偷懶自暴自棄的性格，讓咱們向上精進的鬥志由此而被激發，不是諍友諫友良師益友，而是什麼呢？
                本來，高手不在他用的工具較一般人優利，而是他能善用一切有用可用的工具，如一葦即能渡江，木劍也勝刀劍，化腐朽為神奇，太極而無極，無來亦無去……。感恩感恩　南無阿彌陀佛
                https://ctext.org/library.pl?if=gb&file=78287&page=128#%E8%BA%AB%E5%9C%A8%E6%9C%80%E9%AB%98%E5%B1%A4
                為什麼不這樣想呢？：功夫再好，再加上會用子彈呢？如虎添翼，是不是這麼說的呢？轉個念頭，世界不同，心能轉物，則同如來。我們不用抵抗，而是要內化為我的功夫之一，否則就像滿清末內，維新變法，最後只能亡國絕祀了。感恩感恩 南無阿彌陀佛
            */
            // 建立連接字串
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=查字.mdb";

            // 建立連接物件
            OleDbConnection conn = new OleDbConnection(connectionString);

            // 建立命令物件
            OleDbCommand cmd = conn.CreateCommand();

            // 設定命令的文字
            cmd.CommandText = "SELECT COUNT(*) FROM 異體字反正 WHERE 字 = @word";

            // 設定命令的參數
            cmd.Parameters.AddWithValue("@word", "你要檢查的字");

            // 開啟資料庫連接
            conn.Open();

            // 執行命令並取得結果
            int count = (int)cmd.ExecuteScalar();

            // 關閉資料庫連接
            conn.Close();

            // 判斷結果
            if (count > 0)
            {
                return true;// 資料表中已有該字記錄
            }
            else
            {
                return false;// 資料表中沒有該字記錄
            }

        }
    }
}
