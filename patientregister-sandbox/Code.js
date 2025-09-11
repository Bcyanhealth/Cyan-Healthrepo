/***** CONFIG *****/
const SHEETS = {
	patient: "PatientRegister",
	consult: "ConsultationNotes",
	tele: "Telemedicine",
	payments: "Payments",
	revenue: "RevenueTracker",
	followups: "FollowUps",
	settings: "Settings",
};

const COLS = {
	// zero-based indexes for clarity; adjust if you change layouts
	patient: {
		ts: 0,
		id: 1,
		name: 2,
		email: 3,
		dob: 4,
		age: 5,
		gender: 6,
		phone: 7,
		address: 8,
		reason: 9,
		consent: 10,
		folder: 11,
	},
	consult: {
		ts: 0,
		id: 1,
		email: 2,
		dov: 3,
		cc: 4,
		hpi: 5,
		pmh: 6,
		vitals: 7,
		exam: 8,
		dx1: 9,
		dx2: 10,
		investigations: 11,
		referral: 12,
		clinician: 13,
		signature: 14,
	},
	tele: { ts: 0, name: 1, phone: 2, service: 3, date: 4, time: 5, notes: 6 },
	pay: {
		ts: 0,
		id: 1,
		name: 2,
		service: 3,
		amount: 4,
		method: 5,
		ref: 6,
		by: 7,
	},
	revenue: {
		ts: 0,
		name: 1,
		id: 2,
		date: 3,
		service: 4,
		amount: 5,
		method: 6,
		receipt: 7,
	},
	follow: { ts: 0, name: 1, phone: 2, service: 3, date: 4, time: 5, notes: 6 },
};

function onOpen() {
	SpreadsheetApp.getUi()
		.createMenu("HMIS")
		.addItem("🔧 Check Setup", "checkSetup")
		.addItem("📁 Backfill Patient Folders", "backfillPatientFolders")
		.addSeparator()
		.addItem(
			"📅 Rebuild Follow-up Reminder Trigger",
			"installDailyFollowupTrigger"
		)
		.addToUi();
}

function checkSetup() {
	const s = getSettings();
	const missing = [];
	["ROOT_FOLDER_ID", "CALENDAR_ID"].forEach((k) => {
		if (!s[k]) missing.push(k);
	});
	const msg = missing.length
		? "Missing settings: " + missing.join(", ")
		: "All required settings found. You're good to go!";
	SpreadsheetApp.getUi().alert(msg);
}

function getSettings() {
	const sh = getSheet(SHEETS.settings);
	const rows = sh.getDataRange().getValues().slice(1);
	const map = {};
	rows.forEach((r) => (map[String(r[0]).trim()] = String(r[1] || "").trim()));
	return map;
}

function getSheet(name) {
	return SpreadsheetApp.getActive().getSheetByName(name);
}
function safeGet(v) {
	return v == null ? "" : v;
}

/***** 1) INTAKE → create patient folder if missing *****/
// Run on Form Submit for PatientRegister (or call via menu for backfill)
function onPatientIntakeSubmit(e) {
	const sh = getSheet(SHEETS.patient);
	const row = e ? e.range.getRow() : sh.getLastRow();
	const data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
	const c = COLS.patient;
	const patientID = safeGet(data[c.id]);
	const fullName = safeGet(data[c.name]);

	if (!patientID || !fullName) return;

	// Skip if folder already set
	if (safeGet(data[c.folder])) return;

	const settings = getSettings();
	const root = DriveApp.getFolderById(settings.ROOT_FOLDER_ID);
	const folder = root.createFolder(`${fullName} - ${patientID}`);
	sh.getRange(row, c.folder + 1).setValue(folder.getUrl());
}

function backfillPatientFolders() {
	const sh = getSheet(SHEETS.patient);
	const vals = sh.getDataRange().getValues();
	for (let r = 2; r <= vals.length; r++) {
		onPatientIntakeSubmit({ range: sh.getRange(r, 1) }); // simulate per-row
	}
}

/***** 2) CONSULTATION NOTES → log + (optional) drop a doc in patient folder *****/
// Run on Form Submit for ConsultationNotes
function onConsultationSubmit(e) {
	const sh = getSheet(SHEETS.consult);
	const row = e ? e.range.getRow() : sh.getLastRow();
	const data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
	const c = COLS.consult;

	const patientID = safeGet(data[c.id]);
	const fullName = safeGet(data[c.name]);
	const clinician = safeGet(data[c.clinician]);
	const notes = safeGet(data[c.notes]);

	// Find patient folder
	const psh = getSheet(SHEETS.patient);
	const pvals = psh.getDataRange().getValues();
	let folderUrl = "";
	for (let i = 1; i < pvals.length; i++) {
		if (String(pvals[i][COLS.patient.id]) === patientID) {
			folderUrl = String(pvals[i][COLS.patient.folder] || "");
			break;
		}
	}

	// Optional: create a Google Doc snapshot of notes in folder
	if (folderUrl) {
		const folderId = /folders\/([^\/\?]+)/.exec(folderUrl)?.[1];
		if (folderId) {
			const folder = DriveApp.getFolderById(folderId);
			const doc = DocumentApp.create(
				`Consultation_${patientID}_${new Date().toISOString().slice(0, 10)}`
			);
			doc.getBody().appendParagraph(`Patient: ${fullName} (${patientID})`);
			doc.getBody().appendParagraph(`Clinician: ${clinician}`);
			doc.getBody().appendParagraph("Notes:");
			doc.getBody().appendParagraph(notes);
			doc.saveAndClose();
			const file = DriveApp.getFileById(doc.getId());
			folder.addFile(file);
			DriveApp.getRootFolder().removeFile(file); // keep it only in patient folder
			sh.getRange(row, c.files + 1).setValue(file.getUrl());
		}
	}

	// Update/insert into FollowUps sheet (baseline next follow-up in 30 days)
	upsertFollowup(
		patientID,
		fullName,
		"",
		"",
		new Date(),
		addDays(new Date(), 30),
		"Pending"
	);
}

function upsertFollowup(
	id,
	name,
	email,
	phone,
	lastVisitDate,
	nextDate,
	status
) {
	const sh = getSheet(SHEETS.followups);
	const vals = sh.getDataRange().getValues();
	const c = COLS.follow;

	let foundRow = -1;
	for (let r = 1; r < vals.length; r++) {
		if (String(vals[r][c.id]) === id) {
			foundRow = r + 1;
			break;
		}
	}
	const rowData = [
		id,
		name,
		email || "",
		phone || "",
		toDateOnly(lastVisitDate),
		toDateOnly(nextDate),
		status || "Pending",
	];
	if (foundRow > 0) {
		sh.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
	} else {
		sh.appendRow(rowData);
	}
}

function toDateOnly(d) {
	return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function addDays(d, n) {
	const x = new Date(d);
	x.setDate(x.getDate() + n);
	return x;
}

/***** 3) TELEMEDICINE → create Calendar event with Google Meet *****/
// Run on Form Submit for Telemedicine
// NOTE: For guaranteed Meet creation, enable Advanced Google Services: "Calendar API"
function onTelemedicineSubmit(e) {
	const sh = getSheet(SHEETS.tele);
	const row = e ? e.range.getRow() : sh.getLastRow();
	const data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
	const c = COLS.tele;

	const settings = getSettings();
	const calendarId = settings.CALENDAR_ID;
	if (!calendarId) throw new Error("Missing CALENDAR_ID in Settings");

	const patientEmail = safeGet(data[c.email]);
	const title = safeGet(data[c.title] || `Telemedicine with ${data[c.name]}`);
	const start = new Date(data[c.start]);
	const end = new Date(data[c.end]);

	// Create event via Advanced Calendar to request a Meet link
	const requestId = "req-" + Utilities.getUuid();
	const event = {
		summary: title,
		start: { dateTime: start.toISOString() },
		end: { dateTime: end.toISOString() },
		attendees: patientEmail ? [{ email: patientEmail }] : [],
		conferenceData: {
			createRequest: {
				requestId: requestId,
				conferenceSolutionKey: { type: "hangoutsMeet" },
			},
		},
	};

	const created = Calendar.Events.insert(event, calendarId, {
		conferenceDataVersion: 1,
		sendUpdates: "all",
	});
	const eventId = created.id;
	const meetLink =
		created.hangoutLink ||
		created.conferenceData?.entryPoints?.find(
			(e) => e.entryPointType === "video"
		)?.uri ||
		"";

	sh.getRange(row, c.eventId + 1).setValue(eventId);
	sh.getRange(row, c.meet + 1).setValue(meetLink);
}

/***** 4) PAYMENTS → append to RevenueTracker *****/
// Run on Form Submit for Payments
function onPaymentSubmit(e) {
	const paySh = getSheet(SHEETS.payments);
	const row = e ? e.range.getRow() : paySh.getLastRow();
	const data = paySh.getRange(row, 1, 1, paySh.getLastColumn()).getValues()[0];
	const c = COLS.pay;

	const revSh = getSheet(SHEETS.revenue);
	const out = [
		new Date(), // Date
		safeGet(data[c.service]), // Service
		Number(data[c.amount] || 0), // Amount
		safeGet(data[c.method]), // Method (Cash/Card/Mpesa/etc.)
	];
	revSh.appendRow(out);
}

/***** 5) FOLLOW-UPS → daily reminders (email + optional WhatsApp) *****/
function installDailyFollowupTrigger() {
	// Remove old
	ScriptApp.getProjectTriggers().forEach((t) => {
		if (t.getHandlerFunction() === "dailyFollowupSweep")
			ScriptApp.deleteTrigger(t);
	});
	// New daily at 8am
	ScriptApp.newTrigger("dailyFollowupSweep")
		.timeBased()
		.atHour(8)
		.everyDays(1)
		.create();
}

function dailyFollowupSweep() {
	const sh = getSheet(SHEETS.followups);
	const vals = sh.getDataRange().getValues();
	const c = COLS.follow;
	const today = toDateOnly(new Date());

	for (let r = 1; r < vals.length; r++) {
		const row = vals[r];
		const email = String(row[c.email] || "");
		const phone = String(row[c.phone] || "");
		const next = row[c.next] instanceof Date ? toDateOnly(row[c.next]) : null;
		const status = String(row[c.status] || "");

		if (!next || status.toLowerCase() === "done") continue;

		const daysUntil = Math.round((next - today) / (1000 * 60 * 60 * 24));
		if (daysUntil === 3 || daysUntil === 1 || daysUntil === 0) {
			// Send email reminder
			if (email) sendFollowupEmail(email, row[c.name], next);

			// Optional WhatsApp reminder
			if (phone) sendWhatsAppReminder(phone, row[c.name], next);
		}
	}
}

function sendFollowupEmail(to, name, date) {
	const settings = getSettings();
	const subject = `Clinic follow-up reminder`;
	const body = `Dear ${name},

This is a friendly reminder about your upcoming follow-up on ${date.toDateString()}.
If you need to reschedule, just reply to this email.

Best regards,
Clinic Team`;
	const from = settings.FROM_EMAIL || Session.getActiveUser().getEmail();
	GmailApp.sendEmail(to, subject, body, { name: "Clinic", from: from });
}

// Works with WhatsApp Cloud API or Twilio (HTTP). Configure Settings first.
function sendWhatsAppReminder(phone, name, date) {
	const settings = getSettings();
	const url = settings.WHATSAPP_WEBHOOK_URL;
	const token = settings.WHATSAPP_TOKEN;
	if (!url || !token) return; // not configured

	const msg = `Hi ${name}, reminder of your clinic follow-up on ${date.toDateString()}. Reply if you need to reschedule.`;
	const payload = {
		to: phone,
		message: msg,
		// Adjust to your provider schema (e.g., WhatsApp Cloud API requires "messaging_product","to","type","text":{body})
	};

	const params = {
		method: "post",
		contentType: "application/json",
		payload: JSON.stringify(payload),
		headers: { Authorization: "Bearer " + token },
		muteHttpExceptions: true,
	};
	try {
		UrlFetchApp.fetch(url, params);
	} catch (e) {
		Logger.log(e);
	}
}

function doStuff() {
	Logger.log("hi");
  // comment
}
