# AtenaPrint
はがきの宛名印刷用の Google Slide を作成する Google Apps Script です。サンプルは[こちら](https://script.google.com/macros/s/AKfycbx-pa5AcF4_MW_a2W5K9HWNHHgRiiWgL8vEu-skAlQi25T6a4jSvywrVOIruqw_h5bP/exec)。Google Drive 上のファイルの読み書きやスプレッドシート、スライドの読み書きなど様々な権限を要求されますので、心配なら自分のスクリプトに本コードをコピーして利用してください。

なお、差出人の住所には未対応なので、裏に印刷してください。
## 使い方
1. 公開ページのURLにアクセスする
2. 自分の Google Drive 上のスプレッドシート一覧が表示されるので、宛先情報の入っているスプレッドシートを選択する
3. 利用するシートを選択する
4. 郵便番号、氏名、敬称、住所の記載されている列を選択する
5. 作成ボタンを押下すると、Google Drive 上に印刷用の Google Slide が作成される

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
ただし、今のところ上記のAPIは pageSize の設定に対応しておらず、何を指定しても標準サイズのSlideが作成されてしまう[らしい](https://issuetracker.google.com/issues/119321089)。
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
