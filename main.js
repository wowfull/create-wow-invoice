//請求書自動作成のスタンドアロンプロジェクト版

function fillInWowInvoice() {
  //生成対象月をダイアログで確認
  const today = dayjs.dayjs(new Date());
  // sheetNameはCreatePdfOutputとisFilledにもあるので変更を忘れないこと！
  const sheetName = "請求書作成v2"
  let target = today.add(-1, "month");
  let anotherTarget = null;
  let anotherTargetRaw = null;
  const scriptProperties = PropertiesService.getScriptProperties();
  const cloneOriginSheetId = scriptProperties.getProperty('CLONE_ORIGIN_SHEET_ID');



  //振込先などで未入力箇所がないか確認（内部組み込み関数として動作）
  function isFilled() {
    const ui = SpreadsheetApp.getUi();
    const isFilled = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("B6:G11")
      .getBackgrounds()
      .flat()
      .filter(e => e == "#e06666").length;
    if (isFilled != 0) {
      return ui.alert("警告", "振込先、住所、メールアドレスなど必要な項目を入力したか確認してください。\n\nこの警告を無視して続行しますか？", ui.ButtonSet.YES_NO);
    } else {
      return ui.Button.OK;
    }
  }


  // 請求書シート取得
  const invoice = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  try {
    const ui = SpreadsheetApp.getUi();
    //そもそも「請求書作成」シートがないと話にならないので今のうちに判定
    if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
      const res = ui.alert("警告", `「${sheetName}」シートがありません。\nテンプレートをこのスプレッドシートに新しく生成しますか？\nこのエラーに心当たりが無い場合、シート名を変更していないことを確認してください。`, ui.ButtonSet.YES_NO);
      switch (res) {
        case ui.Button.YES:
          SpreadsheetApp
            .openById(cloneOriginSheetId)
            .getSheetByName(sheetName)
            .copyTo(SpreadsheetApp.getActiveSpreadsheet())
            .setName(sheetName);
          break;
        default: return;
      }
    }
    //振込先などで未入力箇所がないか確認
    switch (isFilled()) {
      case ui.Button.NO:
      case ui.Button.CLOSE: return;
    }
    //請求書手動記載事項が初期値でないかも確認（上に行などが増えていないか数か所確認）
    if (!invoice.getRange("A16").getValue == "No." || !invoice.getRange("F4").getValue == "発行日") {
      ui.alert("エラー", `請求書本体のクリア中に位置エラーが発生しました。行を追加・削除などしていないか確認してください。\n解決しない場合、「${sheetName}」シートの名前を適当なものに変更し、新しく請求書テンプレートを複製してください。`, ui.ButtonSet.OK);
      return;
    }
    //生成対象月の確認
    const res1 = ui.alert("請求書生成対象月の確認", `${today.year()}年${today.month() + 1}月${today.date()}日を発行日として、\n【${target.year()}年${target.month() + 1}月授業料請求書】をPDF出力します。よろしいですか？\n違う日付を指定する場合は「いいえ」を、やめる場合はキャンセルを押してください。`, ui.ButtonSet.YES_NO_CANCEL);
    switch (res1) {
      case ui.Button.YES: /*何もしない*/ break;
      case ui.Button.NO:
        /*日付再入力*/
        anotherTarget = ui.prompt("日付の入力", "請求日の年月日を8桁の半角数字で入力してください。\n例えば、20241001と入力すると「2024年10月01日に提出する予定の、2024年9月に行った授業の請求書」ができます。\n【注意】基本的には参考情報を検索する時のみにご使用ください。\n提出期限を過ぎた請求書や未来の日付の請求書の作成に当たり、問題が生じても責任は取りかねます。また、備考欄に未来や過去の日付の請求書である旨が記載されます。", ui.ButtonSet.OK_CANCEL);
        if (anotherTarget == ui.Button.CANCEL || anotherTarget == ui.Button.CLOSE || anotherTarget == "") {
          return;
        } else {
          anotherTarget = String(anotherTarget.getResponseText());
          if (/[0-9]{8}/.test(anotherTarget)) {
            anotherTarget = anotherTarget.slice(0, 4) + "-" + anotherTarget.slice(4, 6) + "-" + anotherTarget.slice(-2);
            anotherTargetRaw = dayjs.dayjs(anotherTarget);
            anotherTarget = dayjs.dayjs(anotherTarget).add(-1, "month")
          } else {
            ui.alert("エラー", "正しい数値形式での入力を確認できませんでした。", ui.ButtonSet.OK);
            return;
          }
        }
        break;
      case ui.Button.CANCEL:
      case ui.Button.CLOSE: /*強制終了*/ return;
    }
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("エラー", `以下の内容でエラーが発生しました。開発環境以外でこのエラーが発生した場合、ダーさんにお知らせ下さい。\n\n${e}`, ui.ButtonSet.OK)
    console.warn(e);
    console.log("開発用テスト出力開始", target.year(), "年", target.month(), "月分");
  }

  console.log(anotherTarget);
  //anotherTargetの中にnull以外が入っていれば、targetをすげ替える
  if (anotherTarget) {
    target = anotherTarget;
  }
  //結果を連想配列に格納
  let result = [];
  //targetの月と一致するものを請求月分の授業として探していく
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  //生徒名のシート（くん、さん、ちゃん、君、様）で終わるシートをすべて取得し、請求月分の授業があるかどうかを調査
  for (const e in sheets) {
    let n = sheets[e].getName();
    if (n.endsWith("くん") || n.endsWith("さん") || n.endsWith("ちゃん") || n.endsWith("君") || n.endsWith("様")) {
      //このシートについて、今月何回授業があったか判定（＝B列に請求対象月の日付がいくらあったか判定）
      let dates = sheets[e].getRange("B3:B").getValues();
      dates = dates.filter(e => e[0]).flat();
      console.log(dates);
      //日付以外が入っている場合削除
      dates = dates.filter(e => Object.prototype.toString.call(e) === '[object Date]')
      //対象年月以外は消去
      dates = dates.filter(e => e.getMonth() == target.month() && e.getFullYear() == target.year());
      console.log(dates);
      //1つ以上日付の配列が残っていれば、名前と個数を挿入
      if (dates.length) {
        result.push({ name: n, dates: dates });
      }
    }
  }
  //resultが25を超えると多すぎるのでエラー
  if (result.length >= 25) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("エラー", "対象データが多すぎます。現在25人（件）を超えて格納することはできません。ダーさんに問い合わせてください。", ui.ButtonSet.OK);
    return;
  }
  //請求書のガワ改変（請求金額影響なし）・日付編
  if (anotherTarget) {
    invoice.getRange("G4").setFormula(`DATE(${anotherTargetRaw.year()},${anotherTargetRaw.month() + 1},${anotherTargetRaw.date()})`);
  } else {
    invoice.getRange("G4").setFormula("TODAY()");
  }
  //結果を出力（シートに書き込み）
  /*const names = []
  for(const e in result){
    names.push([result[e].name]);
  }
  console.log(names);
  */
  //書き込む前に古い内容をクリア
  invoice.getRange("B17:F41").clearContent();
  //単価部分を元に戻す
  invoice.getRange("E17:E41").setFormulaR1C1('IF(ISBLANK(R[0]C[-3]),"",2000)');
  //名前と回数を入れる
  for (const [i, e] of result.entries()) {
    invoice.getRange(17 + i, 2).setValue(`${e.name}授業（${target.month()}）月`);
    invoice.getRange(17 + i, 4).setValue(e.dates.length);
  }
  console.log(result);

  //請求書をアクティブにしちゃう
  invoice.activate();

  //PDF書き出し
  try {
    //URLを生成する
    const ui = SpreadsheetApp.getUi();
    const res2 = ui.alert("完了", "正常に入力が完了しました。\nこのままPDFも出力しますか？\n少人数授業、その他報酬が発生した業務などを記入する必要がある場合は「いいえ」を押してください。", ui.ButtonSet.YES_NO);
    if (res2 == ui.Button.YES) {
      createPdfOutput();
    }
  } catch (e) {
    console.warn(e);
  }
}


/*const getData = () => {
  const values = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
}*/


//PDF書き出しは独立して別に呼び出せるようにする
function createPdfOutput() {
  const sheetName = "請求書作成v2"
  const invoice = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const ui = SpreadsheetApp.getUi();

  try {
    const caller = createPdfOutput.caller.name;

    if (!caller) {
      switch (isFilled()) {
        case ui.Button.NO:
        case ui.Button.CLOSE: return;
      }
    }
  } catch (e) {
    console.warn("unknown caller");
  }

  //const today = dayjs.dayjs(new Date());
  const issued = dayjs.dayjs(invoice.getRange("G4").getValue());
  // 請求書が有効なものかどうか見ている？詳細不明
  const isValid = invoice.getRange("A46").getValue();
  const target = issued.add(-1, "month");

  //URL生成用のIDを生み出す
  const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const invoiceId = invoice.getSheetId();
  const claimant = invoice.getRange("F6").getValue().toString().replace("　", "").replace(" ", "");
  const fileName = `${target.format("YYYY")}年${target.format("MM")}月_${claimant}_請求書`;

  //URLを生成する
  const baseUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?gid=${invoiceId}&exportFormat=pdf&format=pdf`;
  const pdfOptions = "&exportFormat=pdf&format=pdf"
    + "&size=A4" //用紙サイズ (A4)
    + "&portrait=true"  //用紙の向き true: 縦向き / false: 横向き
    + "&fitw=true"  //ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
    + "&top_margin=0.50" //上の余白
    + "&right_margin=0.50" //右の余白
    + "&bottom_margin=0.50" //下の余白
    + "&left_margin=0.50" //左の余白
    + "&horizontal_alignment=CENTER" //水平方向の位置
    + "&vertical_alignment=TOP" //垂直方向の位置
    + "&printtitle=false" //スプレッドシート名の表示有無
    + "&sheetnames=false" //シート名の表示有無
    + "&gridlines=false" //グリッドラインの表示有無
    + "&fzr=false" //固定行の表示有無
    + "&fzc=false" //固定列の表示有無
    + "&printnotes=false" //メモの表示有無;
  const url = baseUrl + pdfOptions;

  console.log(fileName);
  console.log(url);

  const htmlTemp = HtmlService.createTemplateFromFile("dialog");
  htmlTemp.url = url;
  htmlTemp.fileName = fileName;
  htmlTemp.isValid = isValid;
  html = htmlTemp.evaluate();
  ui.showModalDialog(html, "PDFファイルのダウンロード");
}





















