/**
 * RAINYBRAIN YouTube 仕分け × コメント人格生成ボット
 *
 * 【初回セットアップ】
 *   1. GAS エディタ → 「プロジェクトの設定」→「スクリプトプロパティ」で以下を追加:
 *      KEY_1 〜 KEY_5 に Gemini API キーを1つずつ設定する
 *
 *   2. setupTriggers() を1回だけ手動実行してトリガーを設定する
 *
 * 【Spreadsheet 構成】
 *   Sheet "URLs"    : A列=YouTube URL, B列=ステータス, C列=処理日時
 *   Sheet "Results" : URL・タイトル・チャンネル・カテゴリ・タグ・人格名・人格JSON・処理日時
 *   Sheet "Personas": Channel OS 読み込み用 JSON
 */

const GEMINI_MODEL  = "gemini-2.0-flash";
const MAX_PER_RUN   = 30;
const SLEEP_BETWEEN = 600;

// キーは PropertiesService から取得（コードに書かない）
function getGeminiKeys() {
  const props = PropertiesService.getScriptProperties();
  return ["KEY_1","KEY_2","KEY_3","KEY_4","KEY_5"]
    .map(k => props.getProperty(k))
    .filter(Boolean);
}

let _keyIndex = 0;

// ══════════════════════════════
//  メイン実行（トリガーから呼ぶ）
// ══════════════════════════════
function processVideos() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const urlSheet    = getOrCreateSheet(ss, "URLs",    [["YouTube URL","ステータス","処理日時"]]);
  const resultSheet = getOrCreateSheet(ss, "Results", [["URL","タイトル","チャンネル","カテゴリ","タグ","人格名","人格JSON","処理日時"]]);

  const processedUrls = new Set(
    resultSheet.getLastRow() > 1
      ? resultSheet.getRange(2, 1, resultSheet.getLastRow() - 1, 1).getValues().flat().filter(Boolean)
      : []
  );

  const allUrls = urlSheet.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < allUrls.length && count < MAX_PER_RUN; i++) {
    const url    = String(allUrls[i][0] || "").trim();
    const status = String(allUrls[i][1] || "").trim();

    if (!url || !isYouTubeUrl(url)) continue;
    if (processedUrls.has(url) || status === "完了") continue;

    try {
      urlSheet.getRange(i + 1, 2).setValue("処理中");
      urlSheet.getRange(i + 1, 3).setValue(new Date().toLocaleString("ja-JP"));

      const info = fetchYoutubeInfo(url);
      if (!info) throw new Error("YouTube情報を取得できませんでした");

      const result = classifyAndGeneratePersona(url, info.title, info.author_name);

      resultSheet.appendRow([
        url,
        info.title,
        info.author_name,
        result.category,
        (result.tags || []).join(", "),
        result.persona?.name || "",
        JSON.stringify(result.persona || {}),
        new Date().toISOString()
      ]);

      urlSheet.getRange(i + 1, 2).setValue("完了");
      processedUrls.add(url);
      count++;

      Utilities.sleep(SLEEP_BETWEEN);

    } catch (e) {
      urlSheet.getRange(i + 1, 2).setValue("エラー: " + e.message.slice(0, 80));
      Logger.log("Error on " + url + ": " + e.message);
    }
  }

  exportPersonasSheet(ss, resultSheet);
  Logger.log("処理完了: " + count + "件");
}

// ══════════════════════════════
//  YouTube oEmbed でタイトル取得
// ══════════════════════════════
function fetchYoutubeInfo(videoUrl) {
  const url = "https://www.youtube.com/oembed?url=" + encodeURIComponent(videoUrl) + "&format=json";
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText());
  } catch (e) { return null; }
}

// ══════════════════════════════
//  Gemini で分類 + 人格生成
// ══════════════════════════════
function classifyAndGeneratePersona(url, title, channel) {
  const prompt = `YouTube動画を分析して、カテゴリ分類・タグ付け・コメント人格を生成してください。

動画タイトル: ${title}
チャンネル名: ${channel}
URL: ${url}

JSONのみ出力（説明文不要）:
{
  "category": "ゲーム|教育|エンタメ|音楽|料理|スポーツ|テクノロジー|ビジネス|ライフスタイル|哲学|アート|ニュース|その他",
  "tags": ["タグ1","タグ2","タグ3","タグ4","タグ5"],
  "persona": {
    "name": "人格名（8文字以内）",
    "tone": "コメントスタイル（60文字以内）",
    "roleType": "共感勢|質問勢|知識勢|応援勢|ネタ勢|懐疑勢|妄想勢|初心者勢 のいずれか",
    "examples": ["コメント例1","コメント例2","コメント例3","コメント例4","コメント例5"],
    "catchphrase": "口癖（20文字以内）",
    "color": "#16進数カラー",
    "sourceUrl": "${url}",
    "sourceTitle": "${title.replace(/"/g, "")}"
  }
}`;

  const raw   = callGeminiWithRotation(prompt);
  const match = raw.match(/\{[\s\S]*\}/);
  if (!match) throw new Error("JSONの抽出に失敗");
  return JSON.parse(match[0]);
}

// ══════════════════════════════
//  Gemini 呼び出し（キーローテーション）
// ══════════════════════════════
function callGeminiWithRotation(prompt) {
  const keys = getGeminiKeys();
  if (!keys.length) throw new Error("Gemini APIキーが設定されていません（スクリプトプロパティ KEY_1〜KEY_5）");

  let attempts = 0;
  while (attempts < keys.length) {
    const key = keys[_keyIndex % keys.length];
    const url = "https://generativelanguage.googleapis.com/v1beta/models/"
              + GEMINI_MODEL + ":generateContent?key=" + key;

    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.8, maxOutputTokens: 2048 }
      }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();

    if (code === 429) {
      _keyIndex++;
      attempts++;
      Utilities.sleep(1500);
      continue;
    }
    if (code !== 200) {
      throw new Error("Gemini API " + code + ": " + res.getContentText().slice(0, 200));
    }

    const data = JSON.parse(res.getContentText());
    return data.candidates[0].content.parts[0].text;
  }
  throw new Error("全APIキーがレート上限。次の回復（1分後）まで待機します。");
}

// ══════════════════════════════
//  Personas シート更新
// ══════════════════════════════
function exportPersonasSheet(ss, resultSheet) {
  const sheet = getOrCreateSheet(ss, "Personas", [["人格名","カテゴリ","タグ","ソースURL","ソースタイトル","JSON"]]);

  if (resultSheet.getLastRow() < 2) return;
  const results = resultSheet.getRange(2, 1, resultSheet.getLastRow() - 1, 8).getValues();

  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();

  const rows = results
    .filter(r => r[6])
    .map(r => [r[5], r[3], r[4], r[0], r[1], r[6]]);

  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 6).setValues(rows);
}

// ══════════════════════════════
//  ユーティリティ
// ══════════════════════════════
function isYouTubeUrl(url) {
  return /youtube\.com\/watch|youtu\.be\//.test(url);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    sheet.getRange(1, 1, 1, headers[0].length).setFontWeight("bold");
  }
  return sheet;
}

// ══════════════════════════════
//  トリガー設定（初回1回だけ手動実行）
// ══════════════════════════════
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  [1, 7, 13, 19].forEach(hour => {
    ScriptApp.newTrigger("processVideos").timeBased().atHour(hour).everyDays(1).create();
  });
  Logger.log("トリガー設定完了: 1時/7時/13時/19時");
}

// ══════════════════════════════
//  動作テスト（1件だけ処理）
// ══════════════════════════════
function testSingleVideo() {
  const testUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";
  const info = fetchYoutubeInfo(testUrl);
  if (!info) { Logger.log("YouTube情報取得失敗"); return; }
  Logger.log("タイトル: " + info.title);
  const result = classifyAndGeneratePersona(testUrl, info.title, info.author_name);
  Logger.log(JSON.stringify(result, null, 2));
}
