/**
 * Google Calendar API 連携モジュール
 * カレンダーイベントの取得・作成・更新を行う
 */

/**
 * カレンダーイベントを取得する
 * @param {number} days 取得する日数（デフォルト: 7日）
 * @return {Array} イベントの配列
 */
function getCalendarEvents(days = 7) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const startTime = new Date();
    const endTime = new Date(startTime.getTime() + (days * 24 * 60 * 60 * 1000));

    const events = calendar.getEvents(startTime, endTime);

    return events.map(event => {
      return {
        id: event.getId(),
        title: event.getTitle(),
        description: event.getDescription() || '',
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        location: event.getLocation() || '',
        attendees: event.getGuestList().map(guest => guest.getEmail()),
        isAllDay: event.isAllDayEvent(),
        created: event.getDateCreated(),
        creator: event.getCreators()[0] || '',
        status: getEventStatus(event)
      };
    });

  } catch (error) {
    Logger.log(`カレンダーイベント取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 新しいイベントを作成する
 * @param {Object} eventData イベントデータ
 * @return {Object} 作成されたイベント
 */
function createCalendarEvent(eventData) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();

    let event;
    if (eventData.isAllDay) {
      event = calendar.createAllDayEvent(
        eventData.title,
        eventData.startTime,
        {
          description: eventData.description,
          location: eventData.location,
          guests: eventData.attendees?.join(',') || ''
        }
      );
    } else {
      event = calendar.createEvent(
        eventData.title,
        eventData.startTime,
        eventData.endTime,
        {
          description: eventData.description,
          location: eventData.location,
          guests: eventData.attendees?.join(',') || ''
        }
      );
    }

    Logger.log(`イベント作成完了: ${eventData.title}`);
    return {
      id: event.getId(),
      title: event.getTitle(),
      startTime: event.getStartTime(),
      endTime: event.getEndTime()
    };

  } catch (error) {
    Logger.log(`イベント作成エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 既存イベントを更新する
 * @param {string} eventId イベントID
 * @param {Object} updateData 更新データ
 * @return {Object} 更新されたイベント
 */
function updateCalendarEvent(eventId, updateData) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(new Date(2020, 0, 1), new Date(2030, 11, 31));

    const event = events.find(e => e.getId() === eventId);
    if (!event) {
      throw new Error(`イベントが見つかりません: ${eventId}`);
    }

    // 更新可能な項目を更新
    if (updateData.title) event.setTitle(updateData.title);
    if (updateData.description !== undefined) event.setDescription(updateData.description);
    if (updateData.location !== undefined) event.setLocation(updateData.location);
    if (updateData.startTime && updateData.endTime) {
      event.setTime(updateData.startTime, updateData.endTime);
    }

    Logger.log(`イベント更新完了: ${event.getTitle()}`);
    return {
      id: event.getId(),
      title: event.getTitle(),
      startTime: event.getStartTime(),
      endTime: event.getEndTime()
    };

  } catch (error) {
    Logger.log(`イベント更新エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * イベントを削除する
 * @param {string} eventId イベントID
 * @return {boolean} 削除成功フラグ
 */
function deleteCalendarEvent(eventId) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(new Date(2020, 0, 1), new Date(2030, 11, 31));

    const event = events.find(e => e.getId() === eventId);
    if (!event) {
      Logger.log(`削除対象イベントが見つかりません: ${eventId}`);
      return false;
    }

    event.deleteEvent();
    Logger.log(`イベント削除完了: ${event.getTitle()}`);
    return true;

  } catch (error) {
    Logger.log(`イベント削除エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * スケジュールの空き時間を検索する
 * @param {Date} startDate 検索開始日
 * @param {Date} endDate 検索終了日
 * @param {number} durationMinutes 必要な時間（分）
 * @return {Array} 空き時間の配列
 */
function findAvailableTimeSlots(startDate, endDate, durationMinutes) {
  try {
    const events = getCalendarEvents(
      Math.ceil((endDate - startDate) / (24 * 60 * 60 * 1000))
    );

    const businessHours = {
      start: 9, // 9:00
      end: 18   // 18:00
    };

    const availableSlots = [];
    const currentDate = new Date(startDate);

    while (currentDate <= endDate) {
      // 平日のみチェック（土日を除外）
      if (currentDate.getDay() >= 1 && currentDate.getDay() <= 5) {
        const daySlots = findDayAvailableSlots(
          currentDate,
          events,
          businessHours,
          durationMinutes
        );
        availableSlots.push(...daySlots);
      }

      currentDate.setDate(currentDate.getDate() + 1);
    }

    return availableSlots;

  } catch (error) {
    Logger.log(`空き時間検索エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * イベントの状態を取得する
 * @param {CalendarEvent} event カレンダーイベント
 * @return {string} ステータス
 */
function getEventStatus(event) {
  // GASのCalendarEventオブジェクトから状態を判定
  try {
    const myStatus = event.getMyStatus();
    switch (myStatus) {
      case CalendarApp.GuestStatus.OWNER:
        return 'owner';
      case CalendarApp.GuestStatus.YES:
        return 'accepted';
      case CalendarApp.GuestStatus.NO:
        return 'declined';
      case CalendarApp.GuestStatus.MAYBE:
        return 'tentative';
      default:
        return 'unknown';
    }
  } catch (error) {
    return 'unknown';
  }
}

/**
 * 特定日の空き時間を検索する
 * @param {Date} date 対象日
 * @param {Array} events イベント配列
 * @param {Object} businessHours 営業時間
 * @param {number} durationMinutes 必要時間（分）
 * @return {Array} その日の空き時間
 */
function findDayAvailableSlots(date, events, businessHours, durationMinutes) {
  const dayStart = new Date(date);
  dayStart.setHours(businessHours.start, 0, 0, 0);

  const dayEnd = new Date(date);
  dayEnd.setHours(businessHours.end, 0, 0, 0);

  // その日のイベントのみ抽出
  const dayEvents = events.filter(event => {
    const eventStart = new Date(event.startTime);
    const eventEnd = new Date(event.endTime);
    return (eventStart >= dayStart && eventStart < dayEnd) ||
           (eventEnd > dayStart && eventEnd <= dayEnd) ||
           (eventStart < dayStart && eventEnd > dayEnd);
  });

  // イベントを時間順にソート
  dayEvents.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));

  const availableSlots = [];
  let currentTime = dayStart;

  for (const event of dayEvents) {
    const eventStart = new Date(event.startTime);
    const eventEnd = new Date(event.endTime);

    // 現在時刻とイベント開始時刻の間に空きがあるかチェック
    if (eventStart > currentTime) {
      const gapMinutes = (eventStart - currentTime) / (1000 * 60);
      if (gapMinutes >= durationMinutes) {
        availableSlots.push({
          startTime: new Date(currentTime),
          endTime: new Date(eventStart),
          durationMinutes: gapMinutes
        });
      }
    }

    currentTime = eventEnd > currentTime ? eventEnd : currentTime;
  }

  // 最後のイベント後から営業時間終了までの空きをチェック
  if (currentTime < dayEnd) {
    const gapMinutes = (dayEnd - currentTime) / (1000 * 60);
    if (gapMinutes >= durationMinutes) {
      availableSlots.push({
        startTime: new Date(currentTime),
        endTime: new Date(dayEnd),
        durationMinutes: gapMinutes
      });
    }
  }

  return availableSlots;
}