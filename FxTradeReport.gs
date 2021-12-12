const formId = '';
const to_mailaddress = '';
let mail_subject = '';
let mail_body = '';
const spreadsheet = SpreadsheetApp.openById(formId);
const answer_sheet = spreadsheet.getSheets()[3]; // 4枚目のシート「フォームの回答」を取得
const last_row_num = answer_sheet.getLastRow();
const trade_sheet = spreadsheet.getSheets()[1]; // 2枚目のシート「トレード表」を取得
// 注文日付
const order_date = Utilities.formatDate(answer_sheet.getRange(`B${last_row_num}`).getValue(), 'Asia/Tokyo', 'yyyy/MM/dd');
// 注文時間
const order_time = Utilities.formatDate(answer_sheet.getRange(`C${last_row_num}`).getValue(), 'Asia/Tokyo', 'hh:mm');
// 約定日付
const trade_date = Utilities.formatDate(answer_sheet.getRange(`D${last_row_num}`).getValue(), 'Asia/Tokyo', 'yyyy/MM/dd');
// 約定時間
const trade_time = Utilities.formatDate(answer_sheet.getRange(`E${last_row_num}`).getValue(), 'Asia/Tokyo', 'hh:mm');
// 通貨
const currency_pare = answer_sheet.getRange(`F${last_row_num}`).getValue();
// 時間足
const candle_stick = answer_sheet.getRange(`G${last_row_num}`).getValue();
// ポジション
const position = answer_sheet.getRange(`H${last_row_num}`).getValue();
// 順張りor逆張り
const trend_following_or_contrarian_trading = answer_sheet.getRange(`I${last_row_num}`).getValue();
// ロット数
const lot = answer_sheet.getRange(`J${last_row_num}`).getValue();
// 注文価格
const order_price = answer_sheet.getRange(`K${last_row_num}`).getValue();
// 約定価格
const trade_price = answer_sheet.getRange(`L${last_row_num}`).getValue();
// 損益額
const profits_and_losses = answer_sheet.getRange(`M${last_row_num}`).getValue();
// エントリー要素①
const entry_element_1 = answer_sheet.getRange(`N${last_row_num}`).getValue();
// エントリー要素②
const entry_element_2 = answer_sheet.getRange(`O${last_row_num}`).getValue();
// 決済手法
const settlement_method = answer_sheet.getRange(`P${last_row_num}`).getValue();
// 簡易説明
const brief_explanation = answer_sheet.getRange(`Q${last_row_num}`).getValue();
// 良かった点
const good_point = answer_sheet.getRange(`R${last_row_num}`).getValue();
// 改善点
const bad_point = answer_sheet.getRange(`S${last_row_num}`).getValue();
// 次のトレードに活かせるポイント
const next_point = answer_sheet.getRange(`T${last_row_num}`).getValue();
// トレード画像
const trading_images = answer_sheet.getRange(`U${last_row_num}`).getValue();
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

	return pips;
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
	sendEntryReportMail();
	sendDailyReportMail();
}

/**
 * トレード表シートにコピー
 */
function copyToTradeSheet() {
	trade_sheet.getRange(`B${last_row_num}`).setValue(order_date);
	trade_sheet.getRange(`C${last_row_num}`).setValue(order_time);
	trade_sheet.getRange(`D${last_row_num}`).setValue(trade_date);
	trade_sheet.getRange(`E${last_row_num}`).setValue(trade_time);
	trade_sheet.getRange(`F${last_row_num}`).setValue(currency_pare);
	trade_sheet.getRange(`G${last_row_num}`).setValue(candle_stick);
	trade_sheet.getRange(`H${last_row_num}`).setValue(position);
	trade_sheet.getRange(`I${last_row_num}`).setValue(trend_following_or_contrarian_trading);
	trade_sheet.getRange(`K${last_row_num}`).setValue(lot);
	trade_sheet.getRange(`L${last_row_num}`).setValue(order_price);
	trade_sheet.getRange(`M${last_row_num}`).setValue(trade_price);
	trade_sheet.getRange(`O${last_row_num}`).setValue(profits_and_losses);
	trade_sheet.getRange(`R${last_row_num}`).setValue(entry_element_1);
	trade_sheet.getRange(`S${last_row_num}`).setValue(entry_element_2);
	trade_sheet.getRange(`T${last_row_num}`).setValue(brief_explanation);
	trade_sheet.getRange(`U${last_row_num}`).setValue(good_point);
	trade_sheet.getRange(`V${last_row_num}`).setValue(bad_point);
	trade_sheet.getRange(`W${last_row_num}`).setValue(next_point);
	trade_sheet.getRange(`X${last_row_num}`).setValue(trading_images);
}

/**
 * エントリーメール送信
 */
function sendEntryReportMail() {
	// メール件名
	mail_subject = 'FXトレード報告';
	// メール本文
	mail_body = createEntryReportMessage();
	// メール送信オプション
	let options = getOptions(trading_images);
	// メール送信
	MailApp.sendEmail(to_mailaddress, mail_subject, mail_body, options);
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
⑦${pips}pips
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
	const trading_images_array = trading_images.split(', ');
	let attachments = [];
	for(let i = 0; i < trading_images_array.length; i++){
		let trading_image = trading_images_array[i];
		// トリミング
		trading_image = trading_image.replace('https://drive.google.com/open?id=', '');
		//Googleドライブから画像イメージを取得する
		let attachImg = DriveApp.getFileById(trading_image).getBlob();
		attachments.push(attachImg);
	}
	//オプションで添付ファイルを設定する
	options = {
		"attachments":attachments,
	};
	return options;
}

/**
 * 日報メール送信
 */
function sendDailyReportMail() {
	// メール件名
	mail_subject = 'FX日報';
	// メール本文
	mail_body = createDailyReportMessage();
	// メール送信
	MailApp.sendEmail(to_mailaddress, mail_subject, mail_body);
}

/**
 * 日報LINE本文作成
 */
 function createDailyReportMessage() {
	let dailyReportMessage = `
本日の日報になります！

📣${trade_date}日報📣

《トレード詳細》
${currency_pare} ${pips}pips ${lot}ロット

・${getResultOfTrade()}　合計${pips}pips
・当日損益${profits_and_losses}円
・残高${getBalance()}円
・平均ロット${lot}

【振り返り】
${brief_explanation}

◼︎良かった点
${good_point}

◼︎改善点
${bad_point}

◼︎次のトレードに活かせるポイント
${next_point}
`;
	return dailyReportMessage;
}

/**
 * FX口座残高取得
 */
function getBalance() {
	// 残高
	const currency_pare = trade_sheet.getRange(`Q${last_row_num}`).getValue();
	return currency_pare;
}

/**
 * 1日のトレード結果
 */
function getResultOfTrade() {
	let trade_result = '';
	if(pips >= 0) {
		trade_result = '1戦1勝0敗';
	} else {
		trade_result = '1戦0勝1敗';
	}
	return trade_result;
}