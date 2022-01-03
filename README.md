# はがきの宛名印刷用紙を作成するGAS
はがきの宛名印刷用の Google Slide を作成する Google Apps Script です。サンプルは[こちら](https://script.google.com/macros/s/AKfycbx-pa5AcF4_MW_a2W5K9HWNHHgRiiWgL8vEu-skAlQi25T6a4jSvywrVOIruqw_h5bP/exec)。Google Drive 上のファイルの読み書きやスプレッドシート、スライドの読み書きなど様々な権限を要求されますので、心配なら自分のスクリプトに本コードをコピーして利用してください。

宛先の郵便番号枠の位置は[日本郵便の説明ページ](https://www.post.japanpost.jp/zipcode/zipmanual/p05.html)を参考にしています。差出人の郵便番号枠の場所がよくわからないので、差出人住所の設定には対応していません。裏側にでも書いてください。
## 使い方
1. 公開ページのURLにアクセスする
2. 自分の Google Drive 上のスプレッドシート一覧が表示されるので、宛先情報の入っているスプレッドシートを選択する
3. 利用するシートを選択する
4. 郵便番号、氏名、敬称、住所の記載されている列を選択する
5. 作成ボタンを押下すると、Google Drive 上に印刷用の Google Slide が作成される

## はがきサイズの Google Slide の作り方
仕様上は[Slides.Presentations.create](https://developers.google.com/slides/api/samples/presentation#create_a_new_presentation)の引数でpageSizeを指定できるので、これが使えれば楽なはずです。
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
ただし、今のところ上記のAPIは pageSize の設定に対応しておらず、何を指定しても標準サイズのSlideが作成されてしまう[らしい](https://issuetracker.google.com/issues/119321089)ので、

1. PowerPoint など Google Slide と互換のあるアプリで事前にはがきサイズのスライドファイルを作成
2. それを Base64 形式の文字列で保存
3. [Drive.Files.insert](https://developers.google.com/drive/api/v2/reference/files/insert)で MimeType.GOOGLE_SLIDES 形式でファイルを作成する

という手順ではがきサイズの Google Slide を作成します。具体的なコードは以下の通り。
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
PowerPoint ファイルの中身は Office Open XML という規格で作成された XML ファイルを ZIP で固めたものです。スライドサイズは styles.xml 内で、
```xml:styles.xml
<style:page-layout style:name="pageLayout1">
  <style:page-layout-properties fo:page-width="3.9375in" fo:page-height="5.83333in" style:print-orientation="portrait" style:register-truth-ref-style-name=""/>
</style:page-layout>
```
のように指定されていますので、ここの値を変更して ZIP で固めれば任意のサイズのスライドを作成することが可能です・・・が、正直めんどくさいので、GAS がちゃんと pageSize に対応してくれるのを待つ方が早そう。 
