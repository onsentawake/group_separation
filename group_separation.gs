/**********************************************
 * メインの関数
 * 上部でこの関数を実行するように選択する
 *********************************************/
function createMeetingGroups() {
	// スプレッドシートからデータを取得
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	var data = sheet.getDataRange().getValues(); // 配列に格納
	// Logger.log(data);
	// return;
	// カレンダーIDの設定
	var calendarId = 'カレンダーIDを入力'; // 休暇申請カレンダ

	// data から曜日、時間、グループ数を変数にセット
	var dayOfWeek = data[1][4];  // E2(曜日)
	var time = data[1][5]; // F2(時間)
	var numberOfGroups = data[1][6]; // G2(グループ数)
	var customDate = data[3][4]; // E4(日付指定)

	// dataからB列のみを取得して名簿を作成
	var members = data.slice(1).map(function (row) { return row[1]; }).filter(function (data) { return data !== ''; });

	// 出席できるメンバーをリスト化(欠席者を除外)
	var absentMembers = members.filter(function (member, index) {
		return data[index + 1][2] !== "欠席";
	});

	// 過去のグループ履歴を取得
	var historyRange = sheet.getRange(2, 15, sheet.getLastRow() - 1, numberOfGroups); // O列2行目以降
	var history = historyRange.getValues().map(function (row) {
		return row.map(function (cell) {
			return cell ? cell.split(", ") : [];
		});
	});
	// グループメンバー生成の履歴を取得
	var nonEmptyHistory = history.filter(function (row) {
		return row.some(function (cell) {
			// O列以外の列にデータがあると空の配列ができてしまうので余分な空の配列は除外
			return cell.length > 0;
		});
	});
	// 過去2回分の履歴を取得
	var recentHistory = nonEmptyHistory.slice(-2);

	// GoogleMeet のURLリストを取得
	var googleMeetUrlList = data.slice(6, 6 + numberOfGroups).map(function (row) { return row[4]; });

	// グループを作成
	var groups = createAndCheckGroups(absentMembers, numberOfGroups, recentHistory);

	// Googleカレンダーへの登録
	var date = getMeetingDate(dayOfWeek, time, customDate);

	groups.forEach(function (group, index) {
		var groupDescription = group.join("\n") + "\n\n以下のリンクより会議にお入りください。\n";
		// URLが存在すれば追加し、存在しなければ代替のメッセージを追加
		if (googleMeetUrlList[index]) {
			groupDescription += googleMeetUrlList[index];
		} else {
			groupDescription += "会議室が足りません。担当者は会議スペースを用意して共有ください";
		}

		createEvent(calendarId, date, "グループ " + (index + 1), groupDescription);
	});

	// グループ履歴をシートに追加
	var groupRow = groups.map(function (group) { return group.join(", "); });

	// Find first empty row in the history column and write the new group
	for (var row = 1; row <= sheet.getLastRow(); row++) {
		if (sheet.getRange(row, 15).getValue() === '') {
			sheet.getRange(row, 15, 1, groupRow.length).setValues([groupRow]);
			break;
		}
	}
}


/**********************************************
 * グループを作成する
 * メンバーとグループ数を取得してグループを作成
 *********************************************/
function createGroups(members, numGroups) {
	// 初期化
	var groups = [];
	// 指定したグループ数分の空配列をgroups配列の中にネストして作成
	for (var i = 0; i < numGroups; i++) {
		groups.push([]);
	}
	// グループメンバーを決める
	var groupMembers = randomMembers(members);
	// メンバーを順番にグループに追加
	groupMembers.forEach(function (member, index) {
		// 余ったメンバーを追加
		groups[index % numGroups].push(member);
	});
	return groups;
}


/**********************************************
 * ミーティングの日付を取得する
 *********************************************/
function getMeetingDate(dayOfWeek, time, customDate) {
	if (customDate && customDate !== "") {
		var date = new Date(customDate);
		var hours, minutes;
		if (typeof time !== 'string') {
			// タイムスタンプから時間と分を取得
			time = new Date(time);
			hours = time.getHours();
			minutes = time.getMinutes();
		} else {
			// 時間と分を直接取得
			hours = parseInt(time.split(':')[0]);
			minutes = parseInt(time.split(':')[1]);
		}

		date.setHours(hours);
		date.setMinutes(minutes);
		date.setSeconds(0);
		return date;
	} else {
		return getNextMeetingDate(dayOfWeek, time);
	}
}


/**********************************************
 * 次のミーティングの日付を取得する
 *********************************************/
function getNextMeetingDate(dayOfWeek, time) {
	var days = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
	dayOfWeek = dayOfWeek.toLowerCase();
	var dayIndex = days.indexOf(dayOfWeek);

	// 現在の日付と時間を取得
	var now = new Date();
	var currentDayIndex = now.getDay();
	var diff = dayIndex - currentDayIndex;
	// 次のミーティングの日付を取得
	if (diff <= 0) {
		diff += 7;
	}
	now.setHours(time.getHours(), 0, 0, 0);
	return new Date(now.getTime() + diff * 24 * 60 * 60 * 1000);
}


/**********************************************
 * Googleカレンダーにイベントを作成する
 *********************************************/
function createEvent(calendarId, date, title, description) {
	var calendar = CalendarApp.getCalendarById(calendarId);
	calendar.createEvent(title, date, new Date(date.getTime() + 60 * 60 * 1000), {
		description: description
	});
}


/**********************************************
 * 配列をシャッフルする
 *********************************************/
function randomMembers(members) {
	// 人数分の数字が入る
	var currentIndex = members.length
	var temporaryValue
	var randomIndex;
	// 人数分ループを回す
	while (0 !== currentIndex) {
		randomIndex = Math.floor(Math.random() * currentIndex);
		currentIndex -= 1;
		// 仮置き変数
		temporaryValue = members[currentIndex];
		members[currentIndex] = members[randomIndex];
		members[randomIndex] = temporaryValue;
	}
	// ランダムに生成されたメンバーを格納
	return members;
}


/**********************************************
 * 2つのグループが重複しているか確認する
 *********************************************/
function isDuplicated(newGroups, oldGroups) {
	return newGroups.some(function (newGroup) {
		return oldGroups.some(function (oldGroup) {
			// 新しいグループ（newGroup）と既存のグループ（oldGroup）が完全に同じメンバーで構成されているかを確認する
			return arraysEqual(newGroup, oldGroup);
		});
	});
}


/**********************************************
 * 2つの配列が等しいかどうかを判断する
 *********************************************/
function arraysEqual(a, b) {
	if (a.length !== b.length) return false;

	a.sort();
	b.sort();

	for (var i = 0; i < a.length; i++) {
		if (a[i] !== b[i]) return false;
	}

	return true;
}


/**********************************************
 * グループを作成し、既存のグループと重複がないか確認する
 *********************************************/
function createAndCheckGroups(members, numGroups, history) {
	// 変数を定義
	var groups;
	var retryCount = 0;
	var maxCount = 10;
	// 重複が合った場合かつ、試行回数上限以下の場合に処理を繰り返す
	do {
		groups = createGroups(members, numGroups);
		retryCount++;
	} while (isDuplicated(groups, history) && retryCount < maxCount);
	return groups;
}