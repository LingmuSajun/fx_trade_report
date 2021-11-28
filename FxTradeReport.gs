const formId = '';
const spreadsheet = SpreadsheetApp.openById(formId);
const answer_sheet = spreadsheet.getSheets()[3]; // 4枚目のシート「フォームの回答」を取得
const lastRowNum = answer_sheet.getLastRow();
// 注文日付
const order_date = answer_sheet.getRange(`B${lastRowNum}`).getValue();
// 注文時間
const order_time = answer_sheet.getRange(`C${lastRowNum}`).getValue();
// 約定日付
const trade_date = answer_sheet.getRange(`D${lastRowNum}`).getValue();
// 約定時間
const trade_time = answer_sheet.getRange(`E${lastRowNum}`).getValue();
// 通貨
const currency_pare = answer_sheet.getRange(`F${lastRowNum}`).getValue();
// 時間足
const candle_stick = answer_sheet.getRange(`G${lastRowNum}`).getValue();
// ポジション
const position = answer_sheet.getRange(`H${lastRowNum}`).getValue();
// 順張りor逆張り
const trend_following_or_contrarian_trading = answer_sheet.getRange(`I${lastRowNum}`).getValue();
// ロット数
const lot = answer_sheet.getRange(`J${lastRowNum}`).getValue();
// 注文価格
const order_price = answer_sheet.getRange(`K${lastRowNum}`).getValue();
// 約定価格
const trade_price = answer_sheet.getRange(`L${lastRowNum}`).getValue();
// 損益額
const profits_and_losses = answer_sheet.getRange(`M${lastRowNum}`).getValue();
// エントリー要素①
const entry_element_1 = answer_sheet.getRange(`N${lastRowNum}`).getValue();
// エントリー要素②
const entry_element_2 = answer_sheet.getRange(`O${lastRowNum}`).getValue();
// 決済手法
const settlement_method = answer_sheet.getRange(`P${lastRowNum}`).getValue();
// 簡易説明
const brief_explanation = answer_sheet.getRange(`Q${lastRowNum}`).getValue();
// 良かった点
const good_point = answer_sheet.getRange(`R${lastRowNum}`).getValue();
// 改善点
const bad_point = answer_sheet.getRange(`S${lastRowNum}`).getValue();
// 次のトレードに活かせるポイント
const next_point = answer_sheet.getRange(`T${lastRowNum}`).getValue();
// トレード画像
const trading_images = answer_sheet.getRange(`U${lastRowNum}`).getValue();
// 獲得pips
let pips = calcPips();
// エントリー要素一覧
let entryElements = mergeEntryElements();

/**
 * 獲得pips計算
 */
function calcPips() {
	// クロス円フラグ
	const is_cross_jpy = currency_pare.includes('JPY');

	let pips = '';
	// クロス円かつ買い
	if(is_cross_jpy === true && position === '買い') {
		pips = (trade_price * 1000 - order_price * 1000) / 10;
	// クロス円かつ売り
	} else if(is_cross_jpy === true) {
		pips = (order_price * 1000 - trade_price * 1000) / 10;
	// クロス円ではないかつ買い
	} else if(position === '買い') {
		pips = (trade_price * 1000000 - order_price * 1000000) / 100;
	// クロス円ではないかつ売り
	} else {
		pips = (order_price * 1000000 - trade_price * 1000000) / 100;
	}

	return `${pips}pips`;
}

/**
 * エントリー要素を結合
 */
function mergeEntryElements() {
	let entryElements = '';
	if(entry_element_2 == '-') {
		entryElements = entry_element_1;
	} else {
		entryElements = `${entry_element_1}, ${entry_element_2}`;
	}
	return entryElements;
}

/**
 * TODO: 「フォームの回答」シートの内容をいい感じに「トレード表」シートへ移行。メール送信処理
 *
 */
function execute() {
	copyToTradeSheet();
	sendDailyReportMail();
}

function copyToTradeSheet() {

}

/**
 * 日報メール送信
 */
function sendDailyReportMail() {
	// エントリー報告LINE本文作成
	let entryReportMessage = createEntryReportMessage();

	// 以下メール送信処理
	const to_mailaddress = '';
	const mail_subject = 'テスト FXトレード報告';

	options = getOptions(trading_images);
	//Gmail送信
	MailApp.sendEmail(to_mailaddress, mail_subject, entryReportMessage, options);
}

/**
 * エントリー報告LINE本文作成
 */
 function createEntryReportMessage() {
	let entryReportMessage = `
【エントリー報告】
①${currency_pare}
②${position}
③${order_price}
④${order_date} ${order_time}
⑤${trade_price}
⑥${trade_date} ${trade_time}
⑦${pips}
⑧${entryElements}
⑨${settlement_method}

◼︎説明
${brief_explanation}

◼︎良かった点
${good_point}

◼︎改善点
${bad_point}

◼︎次のトレードに活かせるポイント
${next_point}
`;
	return entryReportMessage;
}

/**
 * メール送信時options取得
 */
function getOptions(trading_images) {
	let options = '';
	let attachImg = '';
	// トリミング
	trading_images = trading_images.replace('https://drive.google.com/open?id=', '');
	//Googleドライブから画像イメージを取得する
	attachImg = DriveApp.getFileById(trading_images).getBlob();
	//オプションで添付ファイルを設定する
	options = {
		"attachments":attachImg,
	};
	return options;
}