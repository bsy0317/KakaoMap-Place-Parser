using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using WinHttp;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Address
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] search_data = {"코다리전문점","코다리","코다리찜"}; //검색어 목록
            string[] rect_area = {                  //검색범위(지역설정)
                "888632,1153050,986392,1222170",    //강원속초
                "936872,1085370,1034632,1154490",   //강원삼척
                "929752,1012730,978632,1047290",    //강원태백
                "979112,641870,1027992,676430",     //경주
                "1022272,717990,1046712,735270",    //포항북구
                "1020792,699310,1045232,716590",    //포항중심
                "1015212,568750,1039652,586030",    //울산
                "989182,465510,1001402,474150",     //부산 해운대
                "971412,477830,983632,486470",      //부산 명륜
                "961727,464585,986167,481865",      //부산 서면
                "921247,425925,970127,460485",      //부산 다대포
                "851767,465445,900647,500005",      //경남창원
                "732207,452765,781087,487325",      //경남진주
                "605487,387925,654367,422485",      //전남순천
                "632527,337285,681407,371845",      //전남여수
                "356867,354765,381307,372045",      //전남목포
                "469767,465145,481987,473785",      //광주1
                "453327,465025,465547,473665",      //광주2
                "446997,446235,495877,480795",      //광주3
                "843657,673175,868097,690455",      //대구구암
                "822557,632615,920317,701735",      //대구광역시
                "790517,714815,839397,749375",      //구미시
                "492917,609355,590677,678475",      //전주
                "571157,771795,620037,806355",      //대전광역시
                "581217,855575,630097,890135",      //청주시
                "507857,962015,556737,996575",      //평택
                "501297,1040875,513517,1049515",    //수원특례시
                "480857,1044405,505297,1061685",    //경기화서
                "504777,1021685,553657,1056245",    //경기동탄
                "426057,1146285,474937,1180845",    //서울1
                "502577,1133005,551457,1167565",    //서울2
                "455977,1085005,553737,1154125",    //서울3
                "445977,1085005,543737,1154125",    //서울4
                "446800,1104105,471960,1117805",    //서울5
                "473600,1101845,498760,1115545",    //서울6
                "508580,1104725,533740,1118425",    //서울7
                "518020,1090405,543180,1104105",    //서울8
                "544500,1082825,569660,1096525",    //기타1
                "502720,1168045,527880,1181745",    //기타2
                "433380,1166985,458540,1180685",    //기타3
                "435340,1155165,460500,1168865",    //기타4
                "455940,1143305,481100,1157005",    //기타5
                "479160,1134275,491740,1141125",    //기타6
                "493070,1129530,499360,1132955",    //기타7
                "488985,1127840,495275,1131265",    //기타8
                "486195,1119110,498775,1125960",    //기타9
                "512445,1123560,525025,1130410",    //기타10
                "500185,1125290,512765,1132140",    //기타11
                "506975,1118590,519555,1125440",    //기타12
                "517375,1109640,529955,1116490",    //기타13
                "517885,1075280,543045,1088980",    //기타14
                "458575,1057310,464865,1060735",    //기타15
                "462870,1054405,475450,1061255",    //기타16
                "485720,1045835,498300,1052685",    //기타17
                "495910,1045135,508490,1051985",    //기타18
                "513845,1026550,520135,1029975",    //기타19
                "512270,1009345,524850,1016195",    //기타20
                "520080,916315,532660,923165",      //기타21
                "526940,916785,552100,930485",      //기타22
                "488900,904985,514060,918685",      //기타23
                "603140,849265,653460,876665",      //기타24
                "586240,779845,598820,786695",      //기타25
                "838990,673865,864150,687565",      //기타26
                "844900,658255,857480,665105",      //기타27
                "861740,652765,874320,659615",      //기타28
                "834220,639805,859380,653505",      //기타29
                "881620,643625,906780,657325",      //기타30
                "858300,374545,908620,401945",      //기타31
                "989220,465615,1001800,472465",     //기타32
                "977790,464785,990370,471635",      //기타33
                "961620,463855,974200,470705",      //기타34
                "948520,442675,973680,456375",      //기타35
                "930540,444835,943120,451685",      //기타36
                "362680,-15905,413000,11495",       //기타37
                "314080,-32445,339240,-18745",      //기타38
                "308840,-99805,409480,-45005",      //기타39
                "457120,-23525,507440,3875",        //기타40
                "457120,-23525,507440,3875",        //기타41
                "418840,440055,519480,494855",      //기타42
                "503220,629395,553540,656795",      //기타43
                "479860,674915,505020,688615",      //기타44
                "874860,849995,900020,863695",      //기타45
                "1034020,720635,1046600,727485",    //기타46
                "400080,1104095,601360,1213695",    //기타47
                "445280,1072095,545920,1126895",    //기타48
                "462320,1032495,562960,1087295",    //기타49
                "651580,1208775,676740,1222475",    //기타50
                "690160,1052455,740480,1079855",    //기타51
                "738080,997295,788400,1024695",     //기타52
                "554520,912375,655160,967175",      //기타53
                "-96600,152855,708520,591255",      //기타54
                "350120,581335,752680,800535",      //기타55
                "638120,594455,1040680,813655",     //기타56
                "550920,761255,651560,816055",      //기타57
                "-751720,50695,2468760,1804295",    //전국
            };

            int max_page = 35; //최대 페이지 탐색 깊이
            Dictionary<string, string> address_dic = new System.Collections.Generic.Dictionary<string, string>(); //중복검사 체크용 트리
            Application application = new Application(); //Excel 초기화
            Workbook workbook = application.Workbooks.Add(); //Excel 새 작업 작성
            Worksheet worksheet = workbook.Worksheets.Add(); //Excel 새 시트 작성
            worksheet.Cells[1, 1].Value = "상호명";            //Excel A1
            worksheet.Cells[1, 2].Value = "도로명 주소";       //Excel B1
            worksheet.Cells[1, 3].Value = "지역";              //Excel C1
            worksheet.Cells[1, 4].Value = "우편번호";          //Excel D1
            worksheet.Cells[1, 5].Value = "전화번호";          //Excel F1
            worksheet.Cells[1, 6].Value = "검색키워드";        //Excel G1

            int x = 2; //Excel 열 좌표
            foreach (string keyword in search_data)
            {
                foreach (string rect in rect_area)
                {
                    for (int i = 1; i <= max_page; i++)
                    {
                        string mapData = getList(keyword, i, rect);
                        if (mapData == "{}") continue;
                        JObject json = JObject.Parse(mapData);
                        JToken json_tok = json["place"];

                        foreach (JToken tmp in json_tok)
                        {
                            string name = tmp["name"].ToString();                       //상호명
                            string address = tmp["new_address"].ToString();             //도로명 주소
                            string address_disp = tmp["new_address_disp"].ToString();   //도로명주소 Separate
                            string zip_code = tmp["new_zipcode"].ToString();            //우편번호
                            string tel = tmp["tel"].ToString();                         //전화번호
                            string query = keyword;                                     //검색키워드

                            if (address_dic.ContainsKey(zip_code)) continue;                      //중복등록 검사
                            address_dic.Add(zip_code, "zip_code");                                //DB등록
                            Console.WriteLine("["+x+"]"+name + "|" + address + "|" + tel);        //화면출력

                            worksheet.Cells[x, 1].Value = name;
                            worksheet.Cells[x, 2].Value = address;
                            worksheet.Cells[x, 3].Value = (address == "" ? "" : address_disp.Split("|")[0]); //지역구분
                            worksheet.Cells[x, 4].Value = zip_code;
                            worksheet.Cells[x, 5].Value = tel;
                            worksheet.Cells[x, 6].Value = query;

                            x++;
                        }

                        //Thread.Sleep(500); //수집 딜레이 설정
                    }
                }
            }
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);          //바탕화면 경로
            string path = Path.Combine(desktopPath, "Excel"+DateTime.Now.ToString("HHmmss")+".xlsx");   //파일이름 병합

            worksheet.Columns.AutoFit();    //Excel 너비 자동 조절
            workbook.SaveAs(path);          //Excel 저장 
            workbook.Close();               //Excel 워크북 닫기
            application.Quit();             //Excel 종료
        }
        static public string getList(string keyword, int page, string rect) //keyword=검색어, page=검색할페이지, rect=검색할 지역
        {
            try
            {
                WinHttpRequest wt = new WinHttpRequest();
                wt.Open("GET", "https://search.map.daum.net/mapsearch/map.daum?callback=jQuery18106817993088107732_1621998218686&q=" + Uri.EscapeUriString(keyword) + "&msFlag=S&page=" + page.ToString() + "&mcheck=Y&rect=" + rect + "&sort=0");
                wt.SetRequestHeader("Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7");
                wt.SetRequestHeader("Cookie", "webid=3c902588ac6e41c5a3d0f21ba448ba69; webid_ts=1619752592088; webid_sync=1621998217863");
                wt.SetRequestHeader("Referer", "https://map.kakao.com/");
                wt.SetRequestHeader("Host", "search.map.daum.net");
                wt.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36");
                wt.SetRequestHeader("Accrpt", "*/*");
                wt.Send();

                //JSON 형식에 맞추어 반환
                return wt.ResponseText.Replace("jQuery18106817993088107732_1621998218686 (", "").Replace(")", "");
            }
            catch
            {
                //오류시 빈 JSON 반환
                return "{}";
            }
        }
    }
}
