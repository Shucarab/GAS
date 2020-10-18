function onOpen() {
  // Uiクラスを取得する
  var ui = SpreadsheetApp.getUi();
  // Uiクラスからメニューを作成する           　　　　
  var menu = ui.createMenu('スクリプト実行');
  // メニューにアイテムを追加する               　　　
  menu.addItem('食べログスクレイピング', 'Execute');
  // メニューをUiクラスに追加する
  menu.addToUi();                                        　　　　
}

//MAxPages の値あたり、20店舗取得するよ
function Execute() {
  var tBook = SpreadsheetApp.getActiveSpreadsheet();
  var tSheet = tBook.getSheetByName("食べログ");
  var MaxPages = tSheet.getRange(1,2).getValue();
  var tURL = tSheet.getRange(2,2).getValue();  
  var tGenre = tSheet.getRange(3,3).getValue(); 
  
  scrapingTabelogTokyo(MaxPages,tURL,tGenre);
}

function scrapingTabelogTokyo(MaxPages,tURL,tGenre) {
  //現在のスプレッドシートを取得
  var aBook = SpreadsheetApp.getActiveSpreadsheet();
  //"食べログスクレイピング"という名前のシートを取得
  var aSheet = aBook.getSheetByName("食べログスクレイピング");
////////////////////////////////////////////////////////////////////////////////////////
//  var MaxPages = 20;
////////////////////////////////////////////////////////////////////////////////////////
  if(tGenre === ""){
    var tURL = tURL + "/rstLst/";  
  }else{
    var tURL = tURL + "/rstLst/" + tGenre +"/" ;  
  }


  for (var page = 1; page <= MaxPages; page++) {
////////////////////////////////////////////////////////////////////////////////////////
//    var url = tURL + "/rstLst/" + page;
    var url = tURL + page;    
    //var url = "https://tabelog.com/tokyo/A1302/A130201/rstLst/"+ page;
    //var url = "https://tabelog.com/tokyo/rstLst/1";//TOKYO
////////////////////////////////////////////////////////////////////////////////////////    
    //if(page === 1){
      //東京のお店(一覧ページの１ページ目)の店名と店ごとのページURLを取得
      //食べログで東京のページのURLを変数urlに代入
      //var url = "https://tabelog.com/tokyo";
    //}
    //URLのページを取得
    var response = UrlFetchApp.fetch(url);
    //HTML文を取得
    var html = response.getContentText('utf-8');
    
    //店名と店ごとのページURLを取得するために正規表現を定義
    var myRegexpParts = /cpy-rst-name.*\/a>/g;
    var myRegexpNamesParts = />.*</g;
    var myRegexpNames = /^.|.$/g;
    var myRegexpUrls = /http.*[1-9]/g;
    
    var myRegexpEval = /<div class="list-rst__rate">([\s\S]*?)<\/div>/g;
    var myRegexpRate = /<span class="c-rating__val/g;
    var myRegexpScore = /list-rst__rating-val">([\s\S]*?)</g;
    var myRegexExpense = /<ul class="list-rst__budget">([\s\S]*?)<\/ul>/g
    var myRegexLunchExpense = /list-rst__budget-val cpy-lunch-budget-val">([\s\S]*?)<\/span>/g;    
    var myRegexDinnerExpense = /list-rst__budget-val cpy-dinner-budget-val">([\s\S]*?)<\/span>/g;
    var myRegexGenre = /<span class="list-rst__area-genre cpy-area-genre">([\s\S]*?)<\/span>/g
    
    //店名と店ごとのページURLが入った部分を配列として取得
    var restaurantParts = html.match(myRegexpParts);
    var restaurantEvals = html.match(myRegexpEval);
    var restaurantExpense = html.match(myRegexExpense); 
    var restaurantGenre = html.match(myRegexGenre);

    //文字列を取得し整形
    for (var i = 0; i < restaurantParts.length; i++) {
      //店名の取得
      var restaurantNamesParts = "" + restaurantParts[i].toString().match(myRegexpNamesParts);
      restaurantNamesParts = restaurantNamesParts.replace('<','');
      restaurantNamesParts = restaurantNamesParts.replace('>',''); 
      //店ごとのページURLの取得
      var restaurantUrls = restaurantParts[i].match(myRegexpUrls);
      //取得した店名をスプレッドシートA列に出力。
      aSheet.getRange(i+2 + (page-1)*20, 1).setValue(restaurantNamesParts);
      //取得した店ごとのページURLをスプレッドシートB列に出力。
      aSheet.getRange(i+2 + (page-1)*20, 2).setValue(restaurantUrls);
 
    }
    //文字列を取得し整形
    for (var i = 0; i < restaurantEvals.length; i++) {
      //Scoreの取得
      if(restaurantEvals[i].match(myRegexpRate)){
        var restaurantEvalsParts = restaurantEvals[i].match(myRegexpScore);
        restaurantEvalsParts = restaurantEvalsParts.toString().match(myRegexpNamesParts);
        restaurantEvalsParts = restaurantEvalsParts.toString().replace('<','');
        restaurantEvalsParts = restaurantEvalsParts.toString().replace('>','');
              
      }else{
        var restaurantEvalsParts = "N/A";
      }

      //取得した店名をスプレッドシートC列に出力。
      aSheet.getRange(i+2 + (page-1)*20, 3).setValue(restaurantEvalsParts);
    }

    //文字列を取得し整形
    for (var i = 0; i < restaurantEvals.length; i++) {
      if(restaurantExpense[i].match(myRegexExpense)){        
        var restaurantLunchExpense = restaurantExpense[i].match(myRegexLunchExpense);
  
        restaurantLunchExpense = restaurantLunchExpense.toString().match(myRegexpNamesParts);
        restaurantLunchExpense = restaurantLunchExpense.toString().replace('<','');
        restaurantLunchExpense = restaurantLunchExpense.toString().replace('>',''); 
    
        var restaurantDinnerExpense = restaurantExpense[i].match(myRegexDinnerExpense);
        restaurantDinnerExpense = restaurantDinnerExpense.toString().match(myRegexpNamesParts);
        restaurantDinnerExpense = restaurantDinnerExpense.toString().replace('<','');
        restaurantDinnerExpense = restaurantDinnerExpense.toString().replace('>','');
          
      }else{

        var restaurantLunchExpense = "N/A";
        var restaurantDinnerExpense = "N/A";
      }
      
      if(restaurantGenre[i].match(myRegexGenre)){        
        var restaurantGenreDetails = restaurantGenre[i].match(myRegexGenre);
  
        restaurantGenreDetails = restaurantGenreDetails.toString().match(myRegexpNamesParts);
        restaurantGenreDetails = restaurantGenreDetails.toString().replace('<','');
        restaurantGenreDetails = restaurantGenreDetails.toString().replace('>',''); 
        restaurantGenreDetails = restaurantGenreDetails.toString().replace('strong>',''); 
        restaurantGenreDetails = restaurantGenreDetails.toString().replace('</strong>',''); 
        restaurantGenreDetails = restaurantGenreDetails.toString().replace('<',''); 
            
      }else{
        var restaurantGenreDetails = "N/A";
      }      
      

      //取得した店名をスプレッドシートC列に出力。

      aSheet.getRange(i+2 + (page-1)*20, 4).setValue(restaurantLunchExpense);
      aSheet.getRange(i+2 + (page-1)*20, 5).setValue(restaurantDinnerExpense);
      aSheet.getRange(i+2 + (page-1)*20, 6).setValue(restaurantGenreDetails);      
    } 

  }
}
