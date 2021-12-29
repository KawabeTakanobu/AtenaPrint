# AtenaPrint
はがきの宛名印刷用の Google Slide を作成する Google Apps Script。
自分の Google Drive から宛名情報の記載されたスプレッドシートを選択し、郵便番号、住所、宛名を記載したはがきサイズのGoogle Slide を作成する。

## はがきサイズの Google Slide の作り方
仕様上は[Slides.Presentations.create](https://developers.google.com/slides/api/samples/presentation#create_a_new_presentation)の引数でpageSizeを指定できるので、これを使えれば楽。
```javascript:sample
let presentation = Slides.Presentations.create({
  title: '宛名印刷用ファイル',
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
```
ただし、2021年12月現在、上記のAPIは pageSize の設定に対応しておらず、何を指定しても標準サイズのSlideが作成されてしまいます。
なので、PowerPoint など Google Slide と互換のあるアプリで事前にはがきサイズのスライドファイルを作成しておき、それを Base64 形式の文字列で持っておき
[Drive.Files.insert](https://developers.google.com/drive/api/v2/reference/files/insert)で MimeType.GOOGLE_SLIDES 形式でファイルを作成することにより、はがきサイズの Google Slide を作成します。
```javascript:コード.gs
const fileName ='宛名印刷用ファイル';
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
```
対象となるファイルのBase64形式での文字列は、以下のようなコードで取得が可能です。
```javascript:sample
function tool() {
  const id = ''; // ファイルのID
  const file = DriveApp.getFileById(id);
  const blob = file.getAs('application/vnd.oasis.opendocument.presentation');
  const base64 = Utilities.base64Encode(blob.getBytes());
  Logger.log(file.getName());
  Logger.log(file.getMimeType());
  Logger.log(base64);
  DriveApp.createFile('base64.txt', base64, 'plain/text');
}
```
