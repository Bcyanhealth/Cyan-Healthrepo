/** -----------------------------
 *  Helper: Parse Time Cell
 *  ----------------------------- */
function parseTimeCell(timeCell) {
  if (!timeCell) return {hours: 0, minutes: 0}; // default if empty

  // Case 1: already a Date object
  if (Object.prototype.toString.call(timeCell) === "[object Date]") {
    return {
      hours: timeCell.getHours(),
      minutes: timeCell.getMinutes(),
    };
  }

  // Case 2: numeric serial
  if (typeof timeCell === "number") {
    const date = new Date(Math.round((timeCell - 25569) * 86400 * 1000));
    return {
      hours: date.getUTCHours(),
      minutes: date.getUTCMinutes(),
    };
  }

  // Case 3: string (e.g. "10:30 AM")
  if (typeof timeCell === "string") {
    const [time, meridian] = timeCell.split(" ");
    let [hours, minutes] = time.split(":").map(Number);
    if (meridian && meridian.toLowerCase() === "pm" && hours < 12) hours += 12;
    if (meridian && meridian.toLowerCase() === "am" && hours === 12) hours = 0;
    return {hours, minutes};
  }

  throw new Error("Unsupported time format");
}

/** -----------------------------
 *  Triggered on form submit
 *  ----------------------------- */
function onFormSubmit(e) {
  const sheet  = e.range.getSheet();
  const row    = e.range.getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const dateObj = values[4];  // adjust to your date column index
  const timeVal = values[5];  // adjust to your time column index

  const {hours, minutes} = parseTimeCell(timeVal);  // <-- now recognized

  const start = new Date(dateObj);
  start.setHours(hours, minutes, 0, 0);
  const end = new Date(start.getTime() + 30 * 60 * 1000);

  const calendarId = 'youremail@gmail.com';
  const requestId  = 'req-' + Utilities.getUuid();
  const event = {
    summary: `Telemedicine with ${values[1]}`,
    start: { dateTime: start.toISOString() },
    end:   { dateTime: end.toISOString() },
    conferenceData: {
      createRequest: {
        requestId: requestId,
        conferenceSolutionKey: { type: 'hangoutsMeet' },
      },
    },
  };

  const created = Calendar.Events.insert(event, calendarId, {
    conferenceDataVersion: 1,
    sendUpdates: 'all',
  });

  const meetLink =
    created.hangoutLink ||
    created.conferenceData?.entryPoints?.find(ep => ep.entryPointType === 'video')?.uri ||
    '';

  // write back to sheet (adjust columns)
  const lastCol = sheet.getLastColumn();
  sheet.getRange(row, lastCol - 1).setValue(created.id);
  sheet.getRange(row, lastCol).setValue(meetLink);
}
