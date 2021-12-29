/************************
 * Drive API を有効にしておく
*************************/
'use strict';

// HTTP GET 受信時の動作
function doGet(e) {
  if(e.parameter.id) {
    let template = HtmlService.createTemplateFromFile('sheet');
    template.options = {
      id: e.parameter.id
    };
    let htmlOutput = template.evaluate();
    htmlOutput.setTitle('データ選択');
    return htmlOutput;
  }
  let htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
  htmlOutput.setTitle('スプレッドシート一覧');
  return htmlOutput;
}

// SpreadSheet のデータを取得する
function getSheetData(id, sheetName) {
  let sheet = SpreadsheetApp.openById(id).getSheetByName(sheetName);
  let retVal = {
    rowCount: sheet.getLastRow(),
    columnCount: sheet.getLastColumn()
  };
  retVal.values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  return retVal;
}

// 文字列を半角にする
function toHankaku(str) {
  return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
}

// 文字列を全角にする
function toZenkaku(str) {
  return str.replace(/[A-Za-z0-9]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) + 0xFEE0);
  });
}

// 表紙を作成する
function setFrontPage(slide, data, scale) {
  // 郵便番号を半角にする
  let postCode = toHankaku(data['郵便番号']);

  // 数字以外の文字列は削除する
  postCode = postCode.replace(/[^\d]/g, '');

  if(postCode.length != 7) {
    throw new Error('郵便番号の桁数が7桁になっていません');
  }

  // 郵便番号の枠の左の位置
  const postCodeLeft = 100 - 47.7 - 8;
  const postCodeWidth = 5.7;
  const postCodeTop = 12;
  const postCodeHeight = 8;
  const textMargin = 10;
  const postCodeFontSize = postCodeHeight * 0.5 * scale;
  const postCodePos = [
    postCodeLeft, postCodeLeft + 7, postCodeLeft + 14, postCodeLeft + 21.6,
    postCodeLeft + 28.4, postCodeLeft + 35.2, postCodeLeft + 42.0
  ];

  // AutoFitでサイズが調整できるとよいのだが、今はまだ機能していないらしいのでFontSize+少し余白で作成する
  for(let i = 0; i < 7; i++) {
    let shape = slide.insertShape(
      SlidesApp.ShapeType.TEXT_BOX,
      (postCodePos[i] + postCodeWidth / 2)* scale - textMargin -postCodeFontSize / 2,
      (postCodeTop + postCodeHeight / 2)* scale - textMargin -postCodeFontSize / 2,
      postCodeFontSize + textMargin * 2,
      postCodeFontSize + textMargin * 2
    );
    let textRange = shape.getText();
    textRange.setText(postCode.charAt(i));
    textRange.getTextStyle().setFontSize(postCodeFontSize);
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  }

  // 住所を記載する
  let addressBlocks = toZenkaku(data['住所'])
  .replace(/[０-９]/g, function(s){
    const kans = '〇一二三四五六七八九';
    const nums = '０１２３４５６７８９';
    for(let i = 0; i < 10; i++) {
      if(s == nums.charAt(i)) {
        return kans.charAt(i);
      }
    }
    return s;
  })
  .replace(/-|‐|－|―|ー/g,'｜')
  .split(/[\s\u3000]+/);

  // 住所の左位置
  const addressLeft = (100 - 20) * scale;
  const addressTop = 30 * scale;
  const addressBottom = (148 - 10) * scale;
  const addressFontSize = 5 * scale;
  const addressTextMax = 17; // 自動計算したいところだが、余白サイズなどが分からないのでとりあえず手動で数える

  let drawAddress = function(text, left, top, alignment) {
    let shape = slide.insertShape(
      SlidesApp.ShapeType.TEXT_BOX,
      left, top,
      addressFontSize + textMargin * 2,
      addressBottom - addressTop
    );
    let textRange = shape.getText();
    textRange.setText(text.split().join('\r\n'));
    textRange.getTextStyle().setFontSize(addressFontSize);
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER).setLineSpacing(100);
    shape.setContentAlignment(alignment);
  };

  let address = addressBlocks[0];
  let left = addressLeft;
  let alignment = SlidesApp.ContentAlignment.TOP;
  for(let i = 1; i < addressBlocks.length; i++) {
    if(address.length + addressBlocks[i].length + 1 > addressTextMax) {
      drawAddress(address, left, addressTop, alignment);
      left -= addressFontSize + textMargin * 2;
      alignment = SlidesApp.ContentAlignment.BOTTOM;
      address = addressBlocks[i];
    }
    else {
      address += '\u3000' + addressBlocks[i];
    }
  }

  // それでもなお長さが足りない場合は、番地の直前で行を変更する
  // 第三新東京市のように地名に漢数字が入っている場合はもはや手補正しかない
  if(address.length > addressTextMax && /^([^一二三四五六七八九]+)([\s\S]*)$/.test(address)) {
    drawAddress(RegExp.$1, left, addressTop, alignment);
    left -= addressFontSize + textMargin * 2;
    drawAddress(RegExp.$2, left, addressTop, SlidesApp.ContentAlignment.BOTTOM);
  }
  else {
    drawAddress(address, left, addressTop, alignment);
  }

  // 宛名を記載する
  const name = data['氏名'].replace(/\s+/g, '\u3000') + '\u3000' + (data['敬称'] || '様');
  const nameFontSize = 8 * scale;
  const nameTop = (postCodeTop + postCodeHeight) * scale;
  const nameBottom = 148 * scale;
  const nameWidth = nameFontSize + textMargin * 2;
  const nameLeft = 100 / 2 * scale - nameWidth / 2;

  let shape = slide.insertShape(
    SlidesApp.ShapeType.TEXT_BOX,
    nameLeft, nameTop,
    nameWidth,
    nameBottom - nameTop
  );
  let textRange = shape.getText();
  textRange.setText(name.split().join('\r\n'));
  textRange.getTextStyle().setFontSize(nameFontSize);
  textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER).setLineSpacing(100);
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
}

// 宛名印刷用のスライドを作成する
function createAtenaSlide(data) {
  // pageSize パラメータは今は効いていないので、今は使用不能
  /*
  let presentation = Slides.Presentations.create({
    title: '宛名印刷用ファイル_' + (function(d){
      return d.getFullYear() + ('00' + (d.getMonth() + 1)).slice(-2) + ('00' + d.getDate()).slice(-2) + ('00' + d.getHours()).slice(-2) + ('00' + d.getMinutes()).slice(-2) + ('00' + d.getSeconds()).slice(-2)
    })(new Date()),
    locale: 'ja',
    pageSize: {
      width: {
        magnitude: 283.5,
        unit: 'PT'
      },
      height: {
        magnitude: 419.5,
        unit: 'PT'
      }
    }
  });
  */

  const fileName ='宛名印刷用ファイル_' + (function(d){
    return d.getFullYear() + ('00' + (d.getMonth() + 1)).slice(-2) + ('00' + d.getDate()).slice(-2) + ('00' + d.getHours()).slice(-2) + ('00' + d.getMinutes()).slice(-2) + ('00' + d.getSeconds()).slice(-2)
  })(new Date());
  const mimeType = 'application/vnd.oasis.opendocument.presentation';
  const blob = Utilities.newBlob(
    Utilities.base64Decode(FILE_BASE64),
    mimeType,
    fileName
  );
  const presentation = SlidesApp.openById(Drive.Files.insert({
    title: fileName,
    mimeType: MimeType.GOOGLE_SLIDES
  }, blob).getId());

  const scale = presentation.getPageWidth() / 100;

  // 余分なページをすべて削除する
  let slides = presentation.getSlides();
  for(let i = slides.length - 1; i >= 0; i--) {
    slides[i].remove();
  }

  data.forEach(function(item){
    setFrontPage(presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK), item, scale);
  });

  return true;
}
