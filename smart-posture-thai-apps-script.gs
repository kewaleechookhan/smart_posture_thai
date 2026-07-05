const SHEET_NAMES = [
  "Participants",
  "PreTest",
  "PostTest",
  "FollowUp",
  "Scoring",
  "Dashboard_Data",
];

function setupSmartPostureThai() {
  const ss = SpreadsheetApp.getActive();
  SHEET_NAMES.forEach((name) => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const assessmentHeaders = [
    "Timestamp",
    "ParticipantID",
    "FullName",
    "Age",
    "Gender",
    "Position",
    "Phase",
    "Risk",
    "Level",
    "Pain",
    "Office",
    "MSD",
    "Ergo",
    "Activity",
    "Satisfaction",
    "KPI",
    "RawJSON",
  ];

  ["PreTest", "PostTest", "FollowUp"].forEach((name) => {
    ss.getSheetByName(name).getRange(1, 1, 1, assessmentHeaders.length).setValues([assessmentHeaders]);
  });

  ss.getSheetByName("Participants").getRange(1, 1, 1, 11).setValues([[
    "ParticipantID",
    "FullName",
    "Age",
    "Gender",
    "Position",
    "WorkHours",
    "ComputerHours",
    "Weight",
    "Height",
    "BMI",
    "Timestamp",
  ]]);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActive();
  const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  setupSmartPostureThai();

  ss.getSheetByName("Participants").appendRow([
    data.participantId,
    data.fullName,
    data.age,
    data.gender,
    data.position,
    data.workHours,
    data.computerHours,
    data.weight,
    data.height,
    data.bmi,
    data.timestamp,
  ]);

  const phase = data.phase || "PreTest";
  const target = ss.getSheetByName(phase) || ss.getSheetByName("PreTest");
  target.appendRow([
    data.timestamp,
    data.participantId,
    data.fullName,
    data.age,
    data.gender,
    data.position,
    phase,
    data.risk,
    data.level,
    data.pain,
    data.office,
    data.msdRaw,
    data.ergoRaw,
    data.activityRaw,
    data.satisfactionRaw,
    data.kpi,
    JSON.stringify(data),
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
