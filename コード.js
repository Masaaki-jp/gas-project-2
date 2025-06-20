function registerChocoZAPtoCalendar() {
  // 検索するメールの条件
  const searchSubject = "【予約確定】";
  const searchFrom = "no-reply@sys.chocozap.jp";

  // 検索クエリを作成
  // is:unread を追加して未読メールのみを対象にする
  const searchQuery = `subject:"${searchSubject}" from:"${searchFrom}" is:unread`;

  // 条件に一致する未読メールのスレッドをすべて取得
  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    Logger.log("該当する未読の予約確定メールが見つかりませんでした。");
    return;
  }

  Logger.log(`${threads.length} 件の該当スレッドが見つかりました。処理を開始します。`);

  // 取得したスレッドをループして処理
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();

    // スレッド内のメッセージをループ（通常は1件だが念のため）
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];

      // 未読のメッセージのみ処理（searchQueryで絞っているが念のため確認）
      if (!message.isUnread()) {
        continue; // 既に既読ならスキップ
      }

      const body = message.getBody();
      const subject = message.getSubject();

      // 抽出結果を格納する変数をこのスコープで宣言
      let menu = null;
      let store = null;
      let finalStartTime = null;
      let finalEndTime = null;
      let extractionSuccessful = false;
      let isFallback = false;
      let extractedDateTimeInfo = null; // 抽出した日時文字列を格納（ログ出力用）


      // --- 主要なパターンで本文から情報を抽出 ---
      // 来店日時から開始日時(YYYY/MM/DD HH:MM)と終了日時(YYYY/MM/DD HH:MM)の両方をキャプチャ
      // グループ1: 開始年月日 (YYYY/MM/DD)
      // グループ2: 開始時刻 (HH:MM)
      // グループ3: 終了年月日 (YYYY/MM/DD)
      // グループ4: 終了時刻 (HH:MM)
      const primaryDateTimeMatch = body.match(/来店日時\s*[:：]?\s*\n?\s*(\d{4}\/\d{2}\/\d{2})\s+(\d{2}:\d{2})〜(\d{4}\/\d{2}\/\d{2})\s+(\d{2}:\d{2})/);
      // メニューと店舗名をキャプチャ
      const primaryMenuMatch = body.match(/■ご予約メニュー\s*[:：]?\s*\n?\s*(.+)\s*\n/);
      const primaryStoreMatch = body.match(/■chocoZAP\s+([^\n]+)/);

      if (primaryDateTimeMatch && primaryMenuMatch && primaryStoreMatch) {
        // 主要なパターンでの抽出がすべて成功
        const dateStr = primaryDateTimeMatch[1]; // 開始年月日YYYY/MM/DD
        const startTimeStr = primaryDateTimeMatch[2]; // 開始時刻 HH:MM
        const endDateStr = primaryDateTimeMatch[3]; // 終了年月日YYYY/MM/DD
        const endTimeStr = primaryDateTimeMatch[4]; // 終了時刻 HH:MM

        menu = primaryMenuMatch[1].trim();
        store = primaryStoreMatch[1].trim();

        // Dateオブジェクトの作成 (主要パターン)
        try {
          const [startY, startM, startD] = dateStr.split('/').map(Number);
          const [startHour, startMinute] = startTimeStr.split(':').map(Number);
          const [endY, endM, endD] = endDateStr.split('/').map(Number);
          const [endHour, endMinute] = endTimeStr.split(':').map(Number);

          // 月は Date コンストラクタでは 0 から始まるため - 1 する
          finalStartTime = new Date(startY, startM - 1, startD, startHour, startMinute);
          finalEndTime = new Date(endY, endM - 1, endD, endHour, endMinute); // 終了年月日と時刻で Date オブジェクトを作成

          // 終了時刻が開始時刻より前になっている場合（日付が同じで時刻だけ戻る場合など）、念のため1日加算するチェック
          // ※終了年月日も取れているので基本的には不要なはずだが、保険として残しても良い
          //   ここでは終了年月日を信じてそのまま進めます。

          extractionSuccessful = true;
          extractedDateTimeInfo = `${dateStr} ${startTimeStr}〜${endDateStr} ${endTimeStr}`;
          Logger.log(`主要なパターンで情報抽出成功。`);

        } catch (e) {
          // 日付/時刻文字列のパースに失敗した場合
          Logger.log(`主要パターンでの日付/時刻パースエラー: ${e}`);
          // この場合は主要パターンでの抽出は失敗とみなす
          finalStartTime = null;
          finalEndTime = null;
        }


      } // 主要パターン抽出終了

      // primaryパターンで抽出に失敗し、かつ fallbackの条件が満たされるかチェック
      if (!extractionSuccessful) {
        // --- 主要なパターンでの抽出に失敗した場合のフォールバック処理 ---
        Logger.log("主要なパターンでの情報抽出に失敗しました。代替パターンを試行します。");

        // フォールバック: 件名から日時 (MM/DD HH:MM) を抽出しようとする
        // ※終了時刻は件名にはない前提。50分後と仮定する。
        const fallbackSubjectDateTimeMatch = subject.match(/(\d{2}\/\d{2})\s+(\d{2}:\d{2})〜?/);

        // フォールバックの条件:
        // 1. 件名から日時 (MM/DD HH:MM) が抽出できた AND
        // 2. 本文からメニュー名が抽出できた (primaryMenuMatch が null でない) AND
        // 3. 本文から店舗名が抽出できた (primaryStoreMatch が null でない)
        if (fallbackSubjectDateTimeMatch && primaryMenuMatch && primaryStoreMatch) {
          // フォールバックでの抽出が成功

          isFallback = true; // フォールバックフラグを立てる

          // メール受信日時から年を取得（予約日時は通常同一年か翌年と仮定）
          const emailDateObj = message.getDate();
          const year = emailDateObj.getFullYear(); // 受信年を取得

          // 件名から抽出した月日 (MM/DD) と年を組み合わせる
          const dateStr = `${year}/${fallbackSubjectDateTimeMatch[1]}`; // 開始年月日として使用
          const startTimeStr = fallbackSubjectDateTimeMatch[2]; // HH:MM (開始)

          // メニューと店舗名は主要パターン抽出時に成功したものを使う
          menu = primaryMenuMatch[1].trim();
          store = primaryStoreMatch[1].trim();

          // Dateオブジェクトの作成 (代替パターン)
          try {
            const [startY, startM, startD] = dateStr.split('/').map(Number);
            const [startHour, startMinute] = startTimeStr.split(':').map(Number);

            // 月は Date コンストラクタでは 0 から始まるため - 1 する
            finalStartTime = new Date(startY, startM - 1, startD, startHour, startMinute);
            // 終了時刻は開始時刻の50分後と仮定
            finalEndTime = new Date(finalStartTime.getTime() + 50 * 60 * 1000); // 50分 = 50 * 60秒 * 1000ミリ秒

            extractionSuccessful = true; // フォールバック成功
            extractedDateTimeInfo = `${dateStr} ${startTimeStr}〜(50分後と仮定)`; // ログ出力用の日時情報
            Logger.log("代替パターンで情報抽出成功 (日時: 件名, メニュー/店舗: 本文)。終了時刻は50分後と仮定。");

          } catch (e) {
            // 日付/時刻文字列のパースに失敗した場合
            Logger.log(`代替パターンでの日付/時刻パースエラー: ${e}`);
            finalStartTime = null;
            finalEndTime = null;
            extractionSuccessful = false; // パース失敗のため抽出失敗とする
          }

        } else {
          // --- どのパターンでも情報抽出が完全に成功しなかった場合 ---
          Logger.log("代替パターンでの情報抽出にも失敗しました。");
          let missingInfo = [];
          // 日時は主要or代替のどちらも失敗した場合に不足とする
          if (!primaryDateTimeMatch && !fallbackSubjectDateTimeMatch) missingInfo.push("日時");
          // メニューは主要パターンが失敗した場合に不足とする
          if (!primaryMenuMatch) missingInfo.push("メニュー");
          // 店舗名は主要パターンが失敗した場合に不足とする
          if (!primaryStoreMatch) missingInfo.push("店舗名");

          Logger.log(`情報抽出失敗。不足情報: ${missingInfo.join(', ')}. 件名: ${subject}, 日付: ${message.getDate()}`);
          // このメールの処理をスキップして次のメールへ
          continue; // extractionSuccessful が false のまま次のループへ
        }
      }

      // --- 情報抽出が成功した場合 (主要パターンまたは代替パターン) ---
      // extractionSuccessful が true の場合のみこのブロックが実行される
      if (extractionSuccessful) {
        // この時点では finalStartTime, finalEndTime, menu, store, isFallback が設定済み

        // カレンダーにイベントを作成
        const calendar = CalendarApp.getDefaultCalendar(); // デフォルトカレンダーを使用

        const eventTitle = `${menu} @ ${store}`; // メニュー名 + 店舗名
        const eventDescription = `chocoZAP 予約\nメニュー: ${menu}\n店舗: ${store}\nメール件名: ${subject}\n抽出方法: ${isFallback ? '代替パターン' : '主要パターン'}`; // 抽出方法も追記

        try {
          // finalStartTime と finalEndTime を使用してイベントを作成
          calendar.createEvent(eventTitle, finalStartTime, finalEndTime, { description: eventDescription });

          // ログ出力
          Logger.log(`カレンダーにイベントを登録しました (${isFallback ? '代替パターン' : '主要パターン'})。"${eventTitle}" (${extractedDateTimeInfo})`);

          message.markRead(); // 正常に処理できたら既読にする
          Logger.log("メールを既読にしました。");

        } catch (e) {
          Logger.log(`カレンダーイベント登録中にエラーが発生しました: ${e}`);
          // カレンダー登録エラーが発生した場合、メールは未読のままにしておく
        }
      }
      // extractionSuccessful が false の場合は continue で次のメールへ進む (上記で処理済み)
    } // messages loop end
  } // threads loop end

  Logger.log("予約確定メール処理が完了しました。");
}

function deleteChocoZAPCancelledEvent() {
  // 検索するメールの条件
  const searchSubjectKeyword = "キャンセル"; // 件名に「キャンセル」が含まれるメールを検索
  const searchFrom = "no-reply@sys.chocozap.jp"; // 送信元アドレス

  // 最もシンプルな検索クエリ: 件名にキーワードが含まれる、指定の送信元からの未読メール
  const searchQuery = `subject:${searchSubjectKeyword} from:"${searchFrom}" is:unread`;

  Logger.log(`[DEBUG] 生成されたGmail検索クエリ: ${searchQuery}`); // ユーザーが確認できるよう残す

  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    Logger.log("該当する未読キャンセルメールが見つかりませんでした。");
    return;
  }

  Logger.log(`${threads.length} 件のキャンセルメールが見つかりました。処理を開始します。`);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const message = thread.getMessages()[0]; // スレッドの最初のメッセージを対象

    // 検索クエリで絞っているが、念のため未読か確認（通常は不要だが安全のため）
    if (!message.isUnread()) {
      continue; // 既に既読ならスキップ
    }

    const body = message.getBody();
    const subject = message.getSubject();

    // メール内容から必要な情報を抽出
    // キャンセルメールの形式に合わせてパターンを調整
    // ■来店日時 YYYY/MM/DD HH:MM形式を抽出 (終了時刻はない場合が多い)
    const dateTimeMatch = body.match(/■来店日時\s*[:：]?\s*\n?\s*(\d{4}\/\d{2}\/\d{2})\s+(\d{2}:\d{2})/);
    // ■キャンセルしたご予約メニュー
    const menuMatch = body.match(/■キャンセルしたご予約メニュー\s*[:：]?\s*\n?\s*(.+)\s*\n/);
    // ■chocoZAP 店舗名
    const storeMatch = body.match(/■chocoZAP\s+([^\n]+)/);


    if (dateTimeMatch && menuMatch && storeMatch) {
      const startDateStr = dateTimeMatch[1]; // YYYY/MM/DD
      const startTimeStr = dateTimeMatch[2]; // HH:MM
      const menu = menuMatch[1].trim();
      const store = storeMatch[1].trim();

      try {
        const [startY, startM, startD] = startDateStr.split('/').map(Number);
        const [startHour, startMinute] = startTimeStr.split(':').map(Number);

        // Dateオブジェクトを作成 (カレンダー検索用)
        const eventStartTimeFromEmail = new Date(startY, startM - 1, startD, startHour, startMinute);

        const calendar = CalendarApp.getDefaultCalendar();
        // 登録時と同じタイトル形式を使用
        const eventTitleToSearch = `${menu} @ ${store}`;

        // カレンダーイベント検索の範囲を絞る（正確な時刻を狙うため、検索範囲はタイト）
        // ただし、カレンダーAPIのgetEventsは範囲指定が必須のため、厳密に開始時刻を含む最小限の範囲を指定
        // イベントの開始時刻がミリ秒単位で一致しない可能性も考慮し、検索範囲を少し広げる方が確実かもしれません。
        // 例: 1分前〜1分後
        const searchRangeStart = new Date(eventStartTimeFromEmail.getTime() - 60 * 1000); // 1分前
        const searchRangeEnd = new Date(eventStartTimeFromEmail.getTime() + 60 * 1000);   // 1分後


        Logger.log(`カレンダー検索開始。タイトル: "${eventTitleToSearch}", 範囲: ${searchRangeStart.toLocaleString()} 〜 ${searchRangeEnd.toLocaleString()}`);

        const events = calendar.getEvents(searchRangeStart, searchRangeEnd, { search: eventTitleToSearch });

        let eventDeleted = false;
        if (events.length > 0) {
          Logger.log(`カレンダーでイベントの検索結果が見つかりました。件数: ${events.length}`);
          Logger.log(`期待するイベント開始時刻 (メールから): ${eventStartTimeFromEmail.toLocaleString()}`);

          for (let k = 0; k < events.length; k++) {
            const event = events[k];
            Logger.log(`検索結果イベント ${k+1}: タイトル: "${event.getTitle()}", 開始時刻: ${event.getStartTime().toLocaleString()}, 終了時刻: ${event.getEndTime().toLocaleString()}`);

            // タイトルと開始時刻（ミリ秒まで）が一致するか確認
            if (event.getTitle() === eventTitleToSearch &&
                event.getStartTime().getTime() === eventStartTimeFromEmail.getTime()) {
              event.deleteEvent();
              Logger.log(`カレンダーからイベント "${event.getTitle()}" (開始時刻: ${startDateStr} ${startTimeStr}) を削除しました。`);
              eventDeleted = true;
              // 一致する最初のイベントを削除したらループを抜ける
              break;
            } else {
              Logger.log(`イベントが一致しませんでした (タイトル不一致または開始時刻不一致)。`);
            }
          }
        }

        if (!eventDeleted) {
          Logger.log(`カレンダーにイベント "${eventTitleToSearch}" (開始時刻: ${startDateStr} ${startTimeStr}) が見つからなかったか、一致するイベントがありませんでした。`);
        }

        message.markRead(); // 処理できたら既読にする
        Logger.log("キャンセルメールを既読にしました。");

      } catch (e) {
        Logger.log(`[ERROR] キャンセル処理中のエラー (日付/時刻パースまたはカレンダー処理): ${e}`);
         // エラーが発生した場合、メールは未読のままにしておく
      }
    } else {
      Logger.log(`[WARNING] キャンセルメールから必要な情報を抽出できませんでした。件名: ${subject}, 日付: ${message.getDate()}`);
      // 抽出できなかった場合、メールは未読のままにしておく
    }
  }
  Logger.log("キャンセルメール処理が完了しました。");
}


/**
 * chocoZAPの「前日確認」メールを検索し、既読にする関数です。
 */
function markChocoZAPConfirmationAsRead() {
  // 検索するメールの条件
  // 件名に「【前日確認】」が含まれるメールを検索
  const searchSubject = "【前日確認】";
  // 送信元アドレス
  const searchFrom = "no-reply@sys.chocozap.jp";

  // 検索クエリを作成: 指定件名、指定送信元、未読メール
  const searchQuery = `subject:"${searchSubject}" from:"${searchFrom}" is:unread`;

  Logger.log(`[DEBUG] 生成されたGmail検索クエリ (前日確認): ${searchQuery}`);

  const threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    Logger.log("該当する未読の前日確認メールは見つかりませんでした。");
    return;
  }

  Logger.log(`${threads.length} 件の前日確認メールが見つかりました。既読化を開始します。`);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();

    // スレッド内のメッセージをループ（通常は1件だが念のため）
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];

      // 未読のメッセージのみ処理（searchQueryで絞っているが念のため確認）
      if (!message.isUnread()) {
        continue; // 既に既読ならスキップ
      }

      try {
        message.markRead(); // 既読にする
        Logger.log(`件名「${message.getSubject()}」のメールを既読にしました。`);
      } catch (e) {
        Logger.log(`[ERROR] 前日確認メールの既読化中にエラー: ${e}`);
        // エラーが発生した場合、メールは未読のままにしておく
      }
    }
  }
   Logger.log("前日確認メール処理が完了しました。");
}


/**
 * ChocoZAPの予約メールからカレンダーにイベントを登録し、
 * キャンセルメールからカレンダーのイベントを削除し、
 * 前日確認メールを既読にするメイン関数です。
 * スクリプトエディタでこの関数を実行するか、トリガーを設定して定期的に実行してください。
 */
function main() {
  Logger.log("--- chocoZAP スケジュール連携処理を開始します ---");

  // 予約メールからのカレンダー登録処理を実行
  registerChocoZAPtoCalendar();

  // キャンセルメールからのカレンダー削除処理を実行
  deleteChocoZAPCancelledEvent();

  // 前日確認メールの既読化処理を実行 (新しく追加)
  markChocoZAPConfirmationAsRead();

  Logger.log("--- chocoZAP スケジュール連携処理を終了します ---");
}