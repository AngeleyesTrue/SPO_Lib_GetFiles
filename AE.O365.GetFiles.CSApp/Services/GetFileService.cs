using AE.O365.GetFiles.CSApp.Common.Authorize;
using AE.O365.GetFiles.CSApp.Common.Extensions;
using AE.O365.GetFiles.CSApp.Common.Interfaces;
using ByteSizeLib;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Security;
using System.Text;

namespace AE.O365.GetFiles.CSApp.Services;

public class GetFileService : IService
{
    private readonly ILogger<GetFileService> _logger;

    public GetFileService(ILogger<GetFileService> logger)
    {
        _logger = logger;
    }

	#region // CamlQuery CreateAllFilesQuery() //
	/// <summary>
	/// 전체 파일/폴더 정보를 가져오는 쿼리를 반환한다.
	/// </summary>
	/// <returns>CamlQuery</returns>
	CamlQuery CreateAllFilesQuery()
    {
        var qry = new CamlQuery();
        qry.ViewXml = "<View Scope=\"RecursiveAll\">" +
            "<Query><ViewFields><FieldRef Name='FSObjType' /><FieldRef Name='ID' /></ViewFields></Query>" +
            "<RowLimit Paged=\"TRUE\">5000</RowLimit>" +
            "</View>";
        return qry;
    } 
    #endregion

    #region // string GetText(string strMessage) //
    /// <summary>
    /// 텍스트를 입력 받는다.
    /// </summary>
    /// <param name="strMessage">출력 메시지</param>
    /// <returns>입력 받은 메시지</returns>
    string GetText(string strMessage)
    {
        Console.WriteLine(strMessage);

        string strReadText = Console.ReadLine();
        if (String.IsNullOrEmpty(strReadText))
        {
            while (true)
            {
                Console.WriteLine(strMessage);
                strReadText = Console.ReadLine();
                if (!String.IsNullOrEmpty(strReadText))
                {
                    return strReadText;
                }
            }
        }

        return strReadText;
    }
	#endregion

	#region // string GetPassword(string strMessage) //
	/// <summary>
	/// 텍스트를 입력 받는다.
	/// </summary>
	/// <param name="strMessage">출력 메시지</param>
	/// <returns>입력 받은 메시지</returns>
	string GetPassword(string strMessage)
    {
        Console.WriteLine(strMessage);

        StringBuilder input = new StringBuilder();
        while (true)
        {
            int x = Console.CursorLeft;
            int y = Console.CursorTop;
            ConsoleKeyInfo key = Console.ReadKey(true);
            if (key.Key == ConsoleKey.Enter)
            {
                Console.WriteLine();
                break;
            }
            if (key.Key == ConsoleKey.Backspace && input.Length > 0)
            {
                input.Remove(input.Length - 1, 1);
                Console.SetCursorPosition(x - 1, y);
                Console.Write(" ");
                Console.SetCursorPosition(x - 1, y);
            }
            else if (key.KeyChar < 32 || key.KeyChar > 126)
            {
                Trace.WriteLine("Output suppressed: no key char"); //catch non-printable chars, e.g F1, CursorUp and so ...
            }
            else if (key.Key != ConsoleKey.Backspace)
            {
                input.Append(key.KeyChar);
                Console.Write("*");
            }
        }
        return input.ToString();
    } 
    #endregion

    public void Run()
    {
		Stopwatch stopwatch = new Stopwatch();
		stopwatch.Start();

        string strSiteUrl = GetText("Url을 입력하여 주십시오.");
        string strTenatId = GetText("Tenat ID를 입력하여 주십시오.");
        string strAccountId = GetText("사용자 ID을 입력하여 주십시오.");
        string strAccountPwd = GetPassword("사용자 비밀번호를 입력하여 주십시오.");

        SecureString oAccountPassword = new SecureString();
		string strPassowrd = strAccountPwd;
		foreach (char c in strPassowrd)
		{
			oAccountPassword.AppendChar(c);
		}

		using (var authenticationManager = new AuthenticationManager(strTenatId))
		using (var ctx = authenticationManager.GetContext(new Uri(strSiteUrl), strAccountId, oAccountPassword))
		{
			ctx.ExecutingWebRequest += delegate (object? sender, WebRequestEventArgs e)
			{
				e.WebRequestExecutor.WebRequest.UserAgent = "NONISV|AE|AE.O365.GetFiles/1.0";
			};

			try
			{
				List<object> arrSiteData = new List<object>();
				List<object> arrListData = new List<object>();
				ConcurrentBag<object> arrCSVData = new ConcurrentBag<object>();

				var site = ctx.Site;
				var web = ctx.Web;
				var lists = web.Lists;
				var Libraries = ctx.LoadQuery(lists.Where(l => l.BaseTemplate == 101 || l.BaseTemplate == 700));

				ctx.Load(site);
				ctx.Load(site, s => s.Usage);
				ctx.Load(web);
				ctx.ExecuteQueryWithIncrementalRetry();

				var siteUsage = ByteSize.FromBytes(site.Usage.Storage);

				long iSiteFileCount = 0;
				foreach (var library in Libraries)
				{
					#region // Set Library Info //
					ctx.Load(library);
					ctx.Load(library, l => l.RootFolder);
					ctx.ExecuteQueryWithIncrementalRetry();

					#region // 시스템 라이브러리 혹은 숨겨진 라이브러리 제외 //
					if (library.Hidden)
						continue;

					switch (library.Title)
					{
						case "Add Lib":
						case "Form Templates":
						case "Site Assets":
						case "Style Library":
						case "Teams Wiki Data":
						case "양식 서식 파일":
						case "스타일 라이브러리":
						case "사이트 자산":
							continue;
					}
					#endregion

					arrListData.Add(new
					{
						SiteId = site.Id,
						SiteUrl = site.ServerRelativeUrl,
						SiteUsage_byte = site.Usage.Storage,
						SiteUsage_Mb = siteUsage.MegaBytes,
						SiteUsage_Gb = siteUsage.GigaBytes,
						WebId = web.Id,
						WebUrl = web.ServerRelativeUrl,
						WebTitle = web.Title,
						ListId = library.Id,
						List = library.Title,
						ListUrl = library.RootFolder.ServerRelativeUrl,
						ItemCount = library.ItemCount,
					});
					#endregion

					#region // 라이브러리에서 파일 정보 추출 //
					var query = CreateAllFilesQuery();
					int iCnt = 1;
					int iFileCount = 0;

					do
					{
						var items = library.GetItems(query);
						ctx.Load(items);
						ctx.ExecuteQueryWithIncrementalRetry();

						var data = items.Where(c => c.FileSystemObjectType.Equals(FileSystemObjectType.File));
						iFileCount += data.Count();

						foreach (var file in data)
						{
							if (file != null && file.FieldValues != null && file.FieldValues.Any())
							{
								string strFileName = Convert.ToString(file.FieldValues["FileLeafRef"]);
								string strExtension = strFileName.Substring(strFileName.LastIndexOf(".") + 1, strFileName.Length - strFileName.LastIndexOf(".") - 1);
								strExtension = strExtension.ToLower();

								DateTime dtCreated = Convert.ToDateTime(file.FieldValues["Created"]);
								DateTime dtModified = Convert.ToDateTime(file.FieldValues["Modified"]);
								long iFileSize = Convert.ToInt64(file.FieldValues["File_x0020_Size"]);
								int FileSize_MB = 0;
								var mb = ByteSize.FromBytes(iFileSize).MegaBytes;
								if (mb > 100)
								{
									FileSize_MB = Convert.ToInt32(Math.Truncate(mb / 100d) * 100);
								}
								else if (mb >= 10 && mb < 100)
								{
									FileSize_MB = Convert.ToInt32(Math.Truncate(mb / 10d) * 10);
								}

								var oCSVData = new
								{
									SiteUrl = site.ServerRelativeUrl,
									SiteUsage_byte = site.Usage.Storage,
									SiteUsage_Mb = ByteSize.FromBytes(site.Usage.Storage).MegaBytes,
									WebId = web.Id,
									WebUrl = web.ServerRelativeUrl,
									WebTitle = web.Title,
									ListId = library.Id,
									List = library.Title,
									ListUrl = library.RootFolder.ServerRelativeUrl,
									ItemCount = library.ItemCount,
									FileId = Convert.ToString(file.FieldValues["UniqueId"]),
									FileName = strFileName,
									FileExtension = strExtension,
									FileUrl = Convert.ToString(file.FieldValues["FileRef"]),
									FileVersion = Convert.ToString(file.FieldValues["_UIVersionString"]),
									FileSize = iFileSize,
									FileSize_MB = FileSize_MB,
									Created = dtCreated.ToString("yyyy.MM.dd hh:mm:sss"),
									CreatedMonth = dtCreated.ToString("yyyy.MM"),
									Modified = dtModified.ToString("yyyy.MM.dd hh:mm:sss"),
									ModifiedMonth = dtModified.ToString("yyyy.MM"),
								};
							}
						}

						_logger.LogInformation($"{web.Title} {library.Title} - ({5000 * iCnt}/{iFileCount}) {items.Count()}/{data.Count()}");

						query.ListItemCollectionPosition = items.ListItemCollectionPosition;

						iCnt++;
					}
					while (query.ListItemCollectionPosition != null);
					#endregion

					iSiteFileCount += iFileCount;
				}

				arrCSVData = new ConcurrentBag<object>();

				_logger.LogInformation($"[{stopwatch.Elapsed}] {web.Title} 사이트 확인 완료");
			}
			catch (Exception ex)
			{
				_logger.LogError(ex, ex.Message);
			}
		}
	}
}
