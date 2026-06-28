# 処理結果・画面例（補足）

> README本体の「9．処理結果・画面例」を補足する詳細資料です。Difyの個別ノード設定や、PDFからのテキスト抽出ログなど、本体には載せていない細かい内容をまとめています。

---

## 補足1．PDF OCRのノード

PDFからのテキスト抽出のノードについて図1に示します。
HTTPリクエストでは、PyMuPDFとGoogle Cloud Vision APIを組み合わせて文字・数字の抽出を行うpdf_ocrのエンドポイントを呼び出して処理を実行しております。

また、更にLLMノードでは、ocr_pdfによるテキスト抽出後に、JSON形式で受け取ったテキストから更にテキスト抽出を行い、抽出精度向上に取り組みました。

[![PDFからのテキスト抽出のノード](https://github.com/yuu-gis0745/prototype/raw/main/docs/images/dify_ocr.png)](/yuu-gis0745/prototype/blob/main/docs/images/dify_ocr.png)

**図1 PDFからのテキスト抽出のノード**

---

## 補足2．PDFからのテキスト抽出結果

PDFから抽出した結果と、PDFファイルにおけるエラーログを図2に示します。
PDFは、Excel様式の森林調査票を一度印刷し、再度PDFでスキャンしたものです(図上段)。
また、抽出した結果もログとして確認用に出力できるようになっております(図中段)。
PDFからの抽出の場合も同様に、エラーログを抽出できております(図下段)。

[![PDFから抽出した結果とPDFファイルにおけるエラーログ](https://github.com/yuu-gis0745/prototype/raw/main/docs/images/pdf_ocr_error_log.png)](/yuu-gis0745/prototype/blob/main/docs/images/pdf_ocr_error_log.png)

**図2 PDFから抽出した結果とPDFファイルにおけるエラーログ**

---

[← README本体に戻る](../README.md)
