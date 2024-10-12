/**
 * Googleドキュメント内の見出し、本文、および箇条書きをGoogleスプレッドシートに階層的に出力するGoogle Apps Script。
 * 
 * 使用方法:
 * 1. docUrlとsheetUrlに、それぞれGoogleドキュメントとGoogleスプレッドシートのURLを設定します。
 * 2. スクリプトを実行すると、Googleドキュメントの内容が階層ごとに解析され、
 *    見出しがカラムに、本文が「本文」カラムに、箇条書きが「箇条書き」カラムに格納されます。
 * 
 * 出力形式の例:
 * 見出し1 | 見出し2 | 本文  | 箇条書き
 * --------|---------|-------|---------
 * 見出し1 | 見出し1-1 | 本文1 | 箇条書き1-1
 * 見出し1 | 見出し1-1 | 本文1 | 箇条書き1-2
 * 見出し1 | 見出し1-2 | 本文2 | 箇条書き2-1
 * 
 * 関数の説明:
 * - importGoogleDocToSheetByUrl():
 *    - GoogleドキュメントとスプレッドシートのURLからIDを取得し、ドキュメントの内容をマークダウン形式に変換します。
 *    - 本文と箇条書きも含め、階層ごとにマークダウン形式のテキストを構築します。
 * 
 * - convertMarkdownToSheet(markdownText, sheetId):
 *    - マークダウン形式のテキストを解析し、見出し、本文、箇条書きのカラムに分けてスプレッドシートに出力します。
 * 
 * - extractIdFromUrl(url):
 *    - GoogleドキュメントまたはスプレッドシートのURLからIDを抽出します。
 * 
 * 注意:
 * - Googleドキュメントのアクセス権とGoogleスプレッドシートの編集権限が必要です。
 * - スプレッドシートの内容はスクリプト実行時にクリアされ、新しいデータで更新されますので、必要に応じてバックアップを取ってください。
 */


function importGoogleDocToSpecifiedSheetByUrl() {
  const docUrl = '<GoogleドキュメントのURL>'; // GoogleドキュメントのURLを設定
  const sheetUrl = '<GoogleスプレッドシートのURL>'; // GoogleスプレッドシートのURLを設定

  const docId = extractIdFromUrl(docUrl);
  const sheetId = extractIdFromUrl(sheetUrl);

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  let markdownText = '';
  const headingLevels = {
    [DocumentApp.ParagraphHeading.HEADING1]: '#',
    [DocumentApp.ParagraphHeading.HEADING2]: '##',
    [DocumentApp.ParagraphHeading.HEADING3]: '###',
    [DocumentApp.ParagraphHeading.HEADING4]: '####',
    [DocumentApp.ParagraphHeading.HEADING5]: '#####',
    [DocumentApp.ParagraphHeading.HEADING6]: '######',
  };

  // 段落ごとに見出しと箇条書きをチェックし、マークダウン形式に変換
  paragraphs.forEach(paragraph => {
    const text = paragraph.getText().trim();
    const heading = paragraph.getHeading();

    if (headingLevels[heading]) {
      markdownText += `${headingLevels[heading]} ${text}\n`;
    } else if (paragraph.getType() === DocumentApp.ElementType.LIST_ITEM) {
      const indentLevel = paragraph.getIndentStart() / 18;
      markdownText += `${' '.repeat(indentLevel * 2)}- ${text}\n`;
    } else if (text) {
      markdownText += `DETAIL: ${text}\n`;
    }
  });

  convertMarkdownToSheet(markdownText, sheetId);
}

function convertMarkdownToSheet(markdownText, sheetId) {
  const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  sheet.clear();

  const lines = markdownText.split('\n');
  const headings = [];
  const data = [];

  let currentDetail = null;

  lines.forEach(line => {
    const headingLevel = (line.match(/^#+/) || [''])[0].length;
    if (headingLevel > 0) {
      // 見出しの処理
      headings[headingLevel - 1] = line.replace(/^#+\s*/, '');
      headings.length = headingLevel;
    } else if (line.startsWith("DETAIL: ")) {
      // 本文の処理
      currentDetail = line.replace("DETAIL: ", '');
    } else if (line.trim().startsWith('- ')) {
      // 箇条書きの処理、本文にぶら下げる
      const bullet = line.replace(/^-+\s*/, '');
      if (currentDetail) {
        data.push([...headings, currentDetail, bullet]);
      }
    }
  });

  // カラムの設定とスプレッドシートへのデータ書き込み
  const maxColumns = Math.max(...data.map(row => row.length));
  const headers = Array.from({ length: maxColumns - 2 }, (_, i) => `見出し${i + 1}`).concat(["本文", "箇条書き"]);
  sheet.appendRow(headers);
  data.forEach(row => sheet.appendRow(row.concat(Array(maxColumns - row.length).fill(''))));
}

function extractIdFromUrl(url) {
  const regex = /\/d\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}
