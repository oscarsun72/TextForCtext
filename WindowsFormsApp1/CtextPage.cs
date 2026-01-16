using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace TextForCtext
{
    public class CtextPage
    {

    }
    /// <summary>
    /// 表示 ctext.org 網頁的不同類型
    /// </summary>
    public enum CtextPageType
    {
        /// <summary>
        /// 未知頁面
        /// </summary>
        Unknown,

        /// <summary>
        /// 書籍首頁
        /// 如： https://ctext.org/library.pl?if=gb&amp;res=6885
        /// </summary>        
        LibraryResource,

        /// <summary>
        /// 圖文對照瀏覽頁面
        /// 如：https://ctext.org/library.pl?if=gb&amp;file=76626&amp;page=1
        /// </summary>
        LibraryFile,

        /// <summary>
        /// 圖文對照編輯頁面，quick edit（可編輯 Wiki）
        /// 如: https://ctext.org/library.pl?if=gb&amp;file=76626&amp;page=5&amp;editwiki=713430#editor
        /// </summary>
        LibraryFileEditWiki,

        /// <summary>
        /// Wiki 資源頁（文字版整部書頁，內含各卷篇名列表）
        /// 如: https://ctext.org/wiki.pl?if=gb&amp;res=53207
        /// </summary>
        WikiResource,

        /// <summary>
        /// Wiki 章節頁（文字版各卷篇章節瀏覽頁面）
        /// 如： https://ctext.org/wiki.pl?if=gb&amp;chapter=713430
        /// </summary>
        WikiChapter,

        /// <summary>
        /// Wiki 編輯頁（文字版各卷篇章節編輯頁面）
        /// 如： https://ctext.org/wiki.pl?if=gb&amp;chapter=713430&amp;action=editchapter#36
        /// </summary>
        WikiEditChapter,
        /// <summary>
        /// 建立新的原典維基項目 上傳新資料 Submit a new text，如 https://ctext.org/wiki.pl?if=en&amp;action=new
        /// </summary>
        SubmitNewText,
        /// <summary>
        /// 修改原典後設資料 Edit text metadata 頁面，如 https://ctext.org/wiki.pl?if=gb&amp;res=53207&action=edit 
        /// </summary>
        EditTextMetadata
    }


    public class CtextPageInfo
    {
        public CtextPageType PageType { get; set; }
        public int? ResId { get; set; }
        public int? FileId { get; set; }
        public int? ChapterId { get; set; }
        public int? PageNumber { get; set; }
        public int? EditWikiId { get; set; } // 新增
    }

    public static class CtextPageClassifier
    {
        // 規則表：路徑 → 判斷邏輯
        private static readonly Dictionary<string, Func<NameValueCollection, CtextPageType>> Rules
            = new Dictionary<string, Func<NameValueCollection, CtextPageType>>(StringComparer.OrdinalIgnoreCase)
            {
                ["library.pl"] = query =>
                {
                    if (query["editwiki"] != null) return CtextPageType.LibraryFileEditWiki;
                    if (query["file"] != null) return CtextPageType.LibraryFile;
                    if (query["res"] != null) return CtextPageType.LibraryResource;

                    return CtextPageType.Unknown;
                },
                ["wiki.pl"] = query =>
                {
                    if (query["action"] == "editchapter") return CtextPageType.WikiEditChapter;
                    if (query["action"] == "edit") return CtextPageType.EditTextMetadata;
                    if (query["action"] == "new") return CtextPageType.SubmitNewText;
                    if (query["chapter"] != null) return CtextPageType.WikiChapter;
                    if (query["res"] != null) return CtextPageType.WikiResource;
                    return CtextPageType.Unknown;
                }
            };

        /// <summary>
        /// 反向映射表：類型 → 範例 URL
        /// Examples 的確是「寫死」的範例 URL。它的用途主要是：
        /// 測試：快速檢查 GetPageType 是否能正確分類。
        /// 文件化：讓未來維護者一眼就知道每個 CtextPageType 對應的 URL 格式。
        /// </summary>
        private static readonly Dictionary<CtextPageType, string> Examples
            = new Dictionary<CtextPageType, string>
        {
        { CtextPageType.LibraryResource, "https://ctext.org/library.pl?if=gb&res=6885" },
        { CtextPageType.LibraryFile,    "https://ctext.org/library.pl?if=gb&file=76626&page=1" },
        { CtextPageType.LibraryFileEditWiki,    "https://ctext.org/library.pl?if=gb&file=76626&page=1&editwiki=713430#editor" },
        { CtextPageType.WikiResource,   "https://ctext.org/wiki.pl?if=gb&res=53207" },
        { CtextPageType.WikiChapter,    "https://ctext.org/wiki.pl?if=gb&chapter=713430" },
        { CtextPageType.WikiEditChapter,"https://ctext.org/wiki.pl?if=gb&chapter=713430&action=editchapter#36" },
        { CtextPageType.EditTextMetadata,"https://ctext.org/wiki.pl?if=gb&res=53207&action=edit" },
        { CtextPageType.SubmitNewText,"https://ctext.org/wiki.pl?if=en&action=new" }
        };
        /// <summary>
        /// 我們可以用一個方法 GetExamples()，在裡面呼叫 BuildUrl 來生成範例字典：
        /// </summary>
        /// <returns></returns>
        public static Dictionary<CtextPageType, string> GetExamples()
        {
            return new Dictionary<CtextPageType, string>
    {
        { CtextPageType.LibraryResource, BuildUrl(CtextPageType.LibraryResource, 6885) },
        { CtextPageType.LibraryFile, BuildUrl(CtextPageType.LibraryFile, 76626, 1) },
        { CtextPageType.LibraryFileEditWiki, BuildUrl(CtextPageType.LibraryFileEditWiki, 76626, 1, 713430) },
        { CtextPageType.WikiResource, BuildUrl(CtextPageType.WikiResource, 53207) },
        { CtextPageType.WikiChapter, BuildUrl(CtextPageType.WikiChapter, 713430) },
        { CtextPageType.WikiEditChapter, BuildUrl(CtextPageType.WikiEditChapter, 713430) },
        { CtextPageType.EditTextMetadata, BuildUrl(CtextPageType.EditTextMetadata, 53207) },
        { CtextPageType.SubmitNewText, BuildUrl(CtextPageType.EditTextMetadata, 0) }
    };
        }
        /* 使用方式:
         * var examples = CtextPageClassifier.GetExamples();
         * foreach (var kvp in examples)
         * {
                Console.WriteLine($"{kvp.Key} → {kvp.Value}");
            }
         */


        // 判斷 URL 類型
        public static CtextPageType GetPageType(string url)
        {
            if (string.IsNullOrEmpty(url)) return CtextPageType.Unknown;

            var uri = new Uri(url);
            var path = uri.AbsolutePath.ToLower();
            var fileName = System.IO.Path.GetFileName(path);
            var query = HttpUtility.ParseQueryString(uri.Query);

            if (Rules.TryGetValue(fileName, out var rule))
            {
                return rule(query);
            }

            return CtextPageType.Unknown;
        }

        // 取得範例 URL
        public static string GetExampleUrl(CtextPageType type)
        {
            return Examples.TryGetValue(type, out var url) ? url : string.Empty;
        }

        // 反向生成方法：給定參數 → 拼出 URL
        public static string BuildUrl(CtextPageType type, int id, int page = 1, int? editWikiId = null)
        {
            switch (type)
            {
                case CtextPageType.LibraryResource:
                    return $"https://ctext.org/library.pl?if=gb&res={id}";
                case CtextPageType.LibraryFile:
                    return $"https://ctext.org/library.pl?if=gb&file={id}&page={page}";
                case CtextPageType.LibraryFileEditWiki:
                    if (editWikiId == null)
                        throw new ArgumentException("editWikiId is required for LibraryFileEditWiki");
                    return $"https://ctext.org/library.pl?if=gb&file={id}&page={page}&editwiki={editWikiId}#editor";
                case CtextPageType.WikiResource:
                    return $"https://ctext.org/wiki.pl?if=gb&res={id}";
                case CtextPageType.WikiChapter:
                    return $"https://ctext.org/wiki.pl?if=gb&chapter={id}";
                case CtextPageType.WikiEditChapter:
                    return $"https://ctext.org/wiki.pl?if=gb&chapter={id}&action=editchapter";
                case CtextPageType.EditTextMetadata:
                    return $"https://ctext.org/wiki.pl?if=gb&res={id}&action=edit";
                case CtextPageType.SubmitNewText:
                    return $"https://ctext.org/wiki.pl?if=en&action=new";
                default:
                    return string.Empty;
            }
        }

        //public class CtextPageInfo
        //{
        //    public CtextPageType PageType { get; set; }
        //    public int? ResId { get; set; }
        //    public int? FileId { get; set; }
        //    public int? ChapterId { get; set; }
        //    public int? PageNumber { get; set; }
        //    public int? EditWikiId { get; set; } // 新增
        //}

        // 🔍 反射式查詢方法：輸入 URL → 拆解出 ID
        public static CtextPageInfo ParseUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return new CtextPageInfo { PageType = CtextPageType.Unknown };

            var uri = new Uri(url);
            var query = HttpUtility.ParseQueryString(uri.Query);

            var info = new CtextPageInfo { PageType = GetPageType(url) };

            if (query["res"] != null && int.TryParse(query["res"], out int resId))
                info.ResId = resId;

            if (query["file"] != null && int.TryParse(query["file"], out int fileId))
                info.FileId = fileId;

            if (query["chapter"] != null && int.TryParse(query["chapter"], out int chapterId))
                info.ChapterId = chapterId;

            if (query["page"] != null && int.TryParse(query["page"], out int pageNum))
                info.PageNumber = pageNum;

            if (query["editwiki"] != null && int.TryParse(query["editwiki"], out int editWikiId))
                info.EditWikiId = editWikiId;

            return info;
        }
    }

}
//使用C#生成XML標記字串 https://copilot.microsoft.com/shares/vVad23KkCDMWSPfRUCRMX https://copilot.microsoft.com/shares/1NUH2riMc7HtAvmVkD6vp