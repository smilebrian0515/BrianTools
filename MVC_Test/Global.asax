<%@ Application Language="C#" %>
<%@ Import Namespace="System.ComponentModel.DataAnnotations" %>
<%@ Import Namespace="System.Web.Routing" %>
<%@ Import Namespace="System.Web.DynamicData" %>
<%@ Import Namespace="System.Web.UI" %>

<script RunAt="server">
    private static MetaModel s_defaultModel = new MetaModel();
    public static MetaModel DefaultModel {
        get {
            return s_defaultModel;
        }
    }

    public static void RegisterRoutes(RouteCollection routes) {
        //                    重要事項: 資料模型註冊 
        // 將此行取消註解，以註冊 ASP.NET 動態資料的 LINQ 到 SQL 模型。
        // 若要設定 ScaffoldAllTables = true:，請務必先確定您希望
        // 資料模型中的所有資料表都支援 Scaffold (即範本) 檢視。若要控制
        // 個別資料表的 Scaffold，請建立資料表的部分類別，並將
        // [ScaffoldTable(true)] 屬性套用至部分類別。
        // 注意: 請確定您在應用程式中將 "YourDataContextType" 變更為資料內容
        // 類別的名稱。
        //DefaultModel.RegisterContext(typeof(YourDataContextType), new ContextConfiguration() { ScaffoldAllTables = false });

        // 下列陳述式支援獨立頁面模式，在這種模式下，List、Detail、Insert 和 
        // Update 等工作都會使用獨立的頁面來執行。若要啟用這個模式，
        // 請取消下列 route 定義的註解，並將後面組合頁面模式區段中的 route 定義標記為註解。
        routes.Add(new DynamicDataRoute("{table}/{action}.aspx") {
            Constraints = new RouteValueDictionary(new { action = "List|Details|Edit|Insert" }),
            Model = DefaultModel
        });

        // 下列陳述式支援 combined-page 模式，在這個模式下，List、Detail、Insert 和
        // Update 等工作都會使用相同的頁面來執行。若要啟用這個模式，
        // 請取消下列 routes 定義的註解，並將上述獨立頁面模式區段中的 route 定義標記為註解。
        //routes.Add(new DynamicDataRoute("{table}/ListDetails.aspx") {
        //    Action = PageAction.List,
        //    ViewName = "ListDetails",
        //    Model = DefaultModel
        //});

        //routes.Add(new DynamicDataRoute("{table}/ListDetails.aspx") {
        //    Action = PageAction.Details,
        //    ViewName = "ListDetails",
        //    Model = DefaultModel
        //});
    }

    private static void RegisterScripts() {
        ScriptManager.ScriptResourceMapping.AddDefinition("jquery", new ScriptResourceDefinition
        {
            Path = "~/Scripts/jquery-1.7.1.min.js",
            DebugPath = "~/Scripts/jquery-1.7.1.js",
            CdnPath = "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.7.1.min.js",
            CdnDebugPath = "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.7.1.js",
            CdnSupportsSecureConnection = true
        });
    }
    
    void Application_Start(object sender, EventArgs e) {
        RegisterRoutes(RouteTable.Routes);
        RegisterScripts();
    }

</script>
