'use strict';

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const CHROME_UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

async function runTests() {
    console.log('====================================================');
    console.log(' STARTING COMPREHENSIVE E2E ATTENDANCE SYSTEM TESTS');
    console.log('====================================================\n');

    const BASE_URL = 'http://127.0.0.1:3000';

    // 1. Check health
    console.log('[TEST 1] Testing /api/health...');
    const healthRes = await fetch(`${BASE_URL}/api/health`);
    const healthData = await healthRes.json();
    if (!healthData.success) throw new Error('Health check failed');
    console.log('  ✓ Server health OK. Status:', healthData.status);

    // 2. Download template and verify 215 roll numbers 25BCS001 - 25BCS215
    console.log('\n[TEST 2] Testing /api/template...');
    const templateRes = await fetch(`${BASE_URL}/api/template`);
    if (!templateRes.ok) throw new Error('Failed to fetch template');
    const templateBuf = await templateRes.arrayBuffer();
    const templateWb = new ExcelJS.Workbook();
    await templateWb.xlsx.load(Buffer.from(templateBuf));
    const templateWs = templateWb.worksheets[0];

    const templateRolls = [];
    for (let r = 2; r <= templateWs.rowCount; r++) {
        const val = templateWs.getRow(r).getCell(1).value;
        if (val) templateRolls.push(String(val).trim());
    }
    if (templateRolls.length !== 215) {
        throw new Error(`Template has ${templateRolls.length} rolls, expected 215`);
    }
    if (templateRolls[0] !== '25BCS001' || templateRolls[214] !== '25BCS215') {
        throw new Error(`Template first/last mismatch: ${templateRolls[0]} - ${templateRolls[214]}`);
    }
    console.log('  ✓ Template verified: 215 students (25BCS001 to 25BCS215).');

    // 3. Testing Upload Validations
    console.log('\n[TEST 3] Testing Invalid Excel Upload Validations...');

    // 3a. Missing 'Roll No.' column
    const badWb1 = new ExcelJS.Workbook();
    const badWs1 = badWb1.addWorksheet('Sheet1');
    badWs1.addRow(['Student Name', 'Department']);
    badWs1.addRow(['Alice', 'CSE']);
    const badBuf1 = await badWb1.xlsx.writeBuffer();

    const badRes1 = await fetch(`${BASE_URL}/api/classes/upload`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            filename: 'Missing_RollNo.xlsx',
            base64: Buffer.from(badBuf1).toString('base64')
        })
    });
    const badData1 = await badRes1.json();
    if (badRes1.status !== 400 || !badData1.error.includes("Roll No.")) {
        throw new Error(`Failed to reject missing Roll No. column: ${JSON.stringify(badData1)}`);
    }
    console.log('  ✓ Correctly rejected missing Roll No. column:', badData1.error);

    // 3b. Unauthorized student (e.g. PhD student 25PCS001)
    const badWb2 = new ExcelJS.Workbook();
    const badWs2 = badWb2.addWorksheet('Sheet1');
    badWs2.addRow(['Roll No.']);
    for (let i = 1; i <= 214; i++) {
        badWs2.addRow(['25BCS' + String(i).padStart(3, '0')]);
    }
    badWs2.addRow(['25PCS001']); // Unauthorized PhD student
    const badBuf2 = await badWb2.xlsx.writeBuffer();

    const badRes2 = await fetch(`${BASE_URL}/api/classes/upload`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            filename: 'Unauthorized_PhD_Student.xlsx',
            base64: Buffer.from(badBuf2).toString('base64')
        })
    });
    const badData2 = await badRes2.json();
    if (badRes2.status !== 400 || !badData2.error.includes('25BCS215')) {
        throw new Error(`Failed to reject unauthorized PhD student: ${JSON.stringify(badData2)}`);
    }
    console.log('  ✓ Correctly rejected unauthorized student:', badData2.error);

    // 3c. Incomplete count (e.g. 210 students)
    const badWb3 = new ExcelJS.Workbook();
    const badWs3 = badWb3.addWorksheet('Sheet1');
    badWs3.addRow(['Roll No.']);
    for (let i = 1; i <= 210; i++) {
        badWs3.addRow(['25BCS' + String(i).padStart(3, '0')]);
    }
    const badBuf3 = await badWb3.xlsx.writeBuffer();

    const badRes3 = await fetch(`${BASE_URL}/api/classes/upload`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            filename: 'Incomplete_Roster.xlsx',
            base64: Buffer.from(badBuf3).toString('base64')
        })
    });
    const badData3 = await badRes3.json();
    if (badRes3.status !== 400 || !badData3.error.includes('Expected 215')) {
        throw new Error(`Failed to reject incomplete roster: ${JSON.stringify(badData3)}`);
    }
    console.log('  ✓ Correctly rejected incomplete student count:', badData3.error);

    // 4. Upload Valid Class 1: Data_Structures.xlsx
    console.log('\n[TEST 4] Uploading Valid Class 1: Data_Structures.xlsx...');
    const dsWb = new ExcelJS.Workbook();
    const dsWs = dsWb.addWorksheet('Attendance');
    dsWs.addRow(['Roll No.']);
    for (let i = 1; i <= 215; i++) {
        dsWs.addRow(['25BCS' + String(i).padStart(3, '0')]);
    }
    const dsBuf = await dsWb.xlsx.writeBuffer();

    const dsRes = await fetch(`${BASE_URL}/api/classes/upload`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            filename: 'Data_Structures.xlsx',
            base64: Buffer.from(dsBuf).toString('base64')
        })
    });
    const dsData = await dsRes.json();
    if (!dsRes.ok || !dsData.success) throw new Error('Data_Structures upload failed: ' + dsData.error);
    const dsFileId = dsData.class.fileId;
    console.log(`  ✓ Data_Structures.xlsx uploaded successfully! FileId: ${dsFileId}`);

    // 5. Upload Valid Class 2: Database_Management.xlsx
    console.log('\n[TEST 5] Uploading Valid Class 2: Database_Management.xlsx...');
    const dbmWb = new ExcelJS.Workbook();
    const dbmWs = dbmWb.addWorksheet('Attendance');
    dbmWs.addRow(['Roll No.']);
    for (let i = 1; i <= 215; i++) {
        dbmWs.addRow(['25BCS' + String(i).padStart(3, '0')]);
    }
    const dbmBuf = await dbmWb.xlsx.writeBuffer();

    const dbmRes = await fetch(`${BASE_URL}/api/classes/upload`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            filename: 'Database_Management.xlsx',
            base64: Buffer.from(dbmBuf).toString('base64')
        })
    });
    const dbmData = await dbmRes.json();
    if (!dbmRes.ok || !dbmData.success) throw new Error('Database_Management upload failed: ' + dbmData.error);
    const dbmFileId = dbmData.class.fileId;
    console.log(`  ✓ Database_Management.xlsx uploaded successfully! FileId: ${dbmFileId}`);

    // 6. Test Class Listing
    console.log('\n[TEST 6] Testing /api/classes listing...');
    const listRes = await fetch(`${BASE_URL}/api/classes`);
    const listData = await listRes.json();
    if (!listData.success || listData.classes.length < 2) {
        throw new Error('Class listing missing uploaded classes');
    }
    console.log(`  ✓ Verified ${listData.classes.length} classes listed.`);

    // 7. Select Data_Structures as active class and start session 1
    console.log('\n[TEST 7] Selecting Data_Structures and starting Session 1...');
    const selRes = await fetch(`${BASE_URL}/api/classes/select`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId: dsFileId })
    });
    const selData = await selRes.json();
    if (!selRes.ok || !selData.success) throw new Error('Failed to select class');
    console.log('  ✓ Class selected:', selData.message);

    const startRes1 = await fetch(`${BASE_URL}/api/reset`, { method: 'POST' });
    const startData1 = await startRes1.json();
    if (!startRes1.ok || !startData1.success) throw new Error('Failed to start session 1');
    const session1Id = startData1.sessionId;
    const session1Col = startData1.sessionColumn;
    console.log(`  ✓ Session 1 started! Column: '${session1Col}', SessionId: ${session1Id}`);

    // Fetch token
    const tokenRes = await fetch(`${BASE_URL}/api/token`);
    const tokenData = await tokenRes.json();
    const token = tokenData.token;
    if (!token) throw new Error('Failed to get token');
    console.log('  ✓ Token obtained:', token);

    // 8. Test Student Attendance Marking
    console.log('\n[TEST 8] Testing Attendance Submissions...');

    // 8a. Valid Student 1: 25BCS001
    const vRes1 = await fetch(`${BASE_URL}/api/validate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S1', 'x-forwarded-for': '10.0.4.1' },
        body: JSON.stringify({ token })
    });
    if (!vRes1.ok) throw new Error('Validation failed for S1: ' + (await vRes1.text()));
    const subRes1 = await fetch(`${BASE_URL}/api/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S1', 'x-forwarded-for': '10.0.4.1' },
        body: JSON.stringify({
            token,
            rollNo: '25BCS001',
            deviceId: 'test-device-student-001'
        })
    });
    const subData1 = await subRes1.json();
    if (!subRes1.ok || !subData1.success) throw new Error('Failed to mark attendance for 25BCS001: ' + subData1.error);
    console.log('  ✓ Attendance successfully recorded for 25BCS001');

    // 8b. Valid Student 2: 25BCS002
    const vRes2 = await fetch(`${BASE_URL}/api/validate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S2', 'x-forwarded-for': '10.0.4.2' },
        body: JSON.stringify({ token })
    });
    if (!vRes2.ok) throw new Error('Validation failed for S2: ' + (await vRes2.text()));
    const subRes2 = await fetch(`${BASE_URL}/api/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S2', 'x-forwarded-for': '10.0.4.2' },
        body: JSON.stringify({
            token,
            rollNo: '25BCS002',
            deviceId: 'test-device-student-002'
        })
    });
    const subData2 = await subRes2.json();
    if (!subRes2.ok || !subData2.success) throw new Error('Failed to mark attendance for 25BCS002: ' + subData2.error);
    console.log('  ✓ Attendance successfully recorded for 25BCS002');

    // 8c. UNAUTHORIZED STUDENT REJECTION (PhD student 25PCS001 / out of roster)
    console.log('\n[TEST 8c] Testing Unauthorized Student Rejection...');
    await fetch(`${BASE_URL}/api/validate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S_PHD', 'x-forwarded-for': '10.0.4.3' },
        body: JSON.stringify({ token })
    });
    const unauthRes = await fetch(`${BASE_URL}/api/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S_PHD', 'x-forwarded-for': '10.0.4.3' },
        body: JSON.stringify({
            token,
            rollNo: '25PCS001', // PhD student
            deviceId: 'test-device-student-phd'
        })
    });
    const unauthData = await unauthRes.json();
    if (unauthRes.status !== 403 || unauthData.error !== 'Attendance Rejected: You are not registered for this class.') {
        throw new Error(`Expected rejection 'Attendance Rejected: You are not registered for this class.', got: ${JSON.stringify(unauthData)}`);
    }
    console.log('  ✓ Correctly rejected unauthorized student with:', unauthData.error);

    // 8d. DUPLICATE ATTENDANCE REJECTION
    console.log('\n[TEST 8d] Testing Duplicate Submission Rejection...');
    await fetch(`${BASE_URL}/api/validate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S_DUP', 'x-forwarded-for': '10.0.4.4' },
        body: JSON.stringify({ token })
    });
    const dupRes = await fetch(`${BASE_URL}/api/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S_DUP', 'x-forwarded-for': '10.0.4.4' },
        body: JSON.stringify({
            token,
            rollNo: '25BCS001', // Already submitted
            deviceId: 'test-device-student-dup'
        })
    });
    const dupData = await dupRes.json();
    if (dupRes.status !== 409) {
        throw new Error(`Expected 409 conflict on duplicate rollNo, got ${dupRes.status}: ${JSON.stringify(dupData)}`);
    }
    console.log('  ✓ Correctly rejected duplicate submission:', dupData.error);

    // 9. Stop Session 1 and verify Excel file
    console.log('\n[TEST 9] Stopping Session 1 and validating updated Excel file...');
    const stopRes1 = await fetch(`${BASE_URL}/api/stop-session`, { method: 'POST' });
    const stopData1 = await stopRes1.json();
    if (!stopRes1.ok || !stopData1.success) throw new Error('Stop session 1 failed');

    // Download updated Excel file
    const dlRes1 = await fetch(`${BASE_URL}/api/download`);
    const dlBuf1 = await dlRes1.arrayBuffer();
    const updatedWb1 = new ExcelJS.Workbook();
    await updatedWb1.xlsx.load(Buffer.from(dlBuf1));
    const wsCheck1 = updatedWb1.worksheets[0];

    console.log(`  ✓ Excel column count after Session 1: ${wsCheck1.columnCount}`);
    console.log(`  ✓ Header column 2 name: '${wsCheck1.getRow(1).getCell(2).value}'`);

    const cell25BCS001_s1 = wsCheck1.getRow(2).getCell(2).value;
    const cell25BCS002_s1 = wsCheck1.getRow(3).getCell(2).value;
    const cell25BCS003_s1 = wsCheck1.getRow(4).getCell(2).value;
    const cell25BCS215_s1 = wsCheck1.getRow(216).getCell(2).value;

    if (cell25BCS001_s1 !== 1 || cell25BCS002_s1 !== 1 || cell25BCS003_s1 !== 0 || cell25BCS215_s1 !== 0) {
        throw new Error(`Excel values mismatch in Session 1: 001=${cell25BCS001_s1}, 002=${cell25BCS002_s1}, 003=${cell25BCS003_s1}, 215=${cell25BCS215_s1}`);
    }
    console.log('  ✓ Verified Session 1 records: 25BCS001=1, 25BCS002=1, 25BCS003=0, 25BCS215=0');

    // 10. Start Session 2 on the SAME day and test non-overwriting
    console.log('\n[TEST 10] Starting Session 2 on same day (Testing Session 2 column creation)...');
    const startRes2 = await fetch(`${BASE_URL}/api/reset`, { method: 'POST' });
    const startData2 = await startRes2.json();
    if (!startRes2.ok || !startData2.success) throw new Error('Failed to start session 2: ' + startData2.error);
    const session2Col = startData2.sessionColumn;
    console.log(`  ✓ Session 2 started! Column: '${session2Col}'`);
    if (!session2Col.includes('(Session 2)')) {
        throw new Error(`Expected column to contain '(Session 2)', got '${session2Col}'`);
    }

    const tokenRes2 = await fetch(`${BASE_URL}/api/token`);
    const tokenData2 = await tokenRes2.json();
    const token2 = tokenData2.token;

    // Student 25BCS003 scans in Session 2
    await fetch(`${BASE_URL}/api/validate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S3', 'x-forwarded-for': '10.0.4.5' },
        body: JSON.stringify({ token: token2 })
    });
    const subRes2_3 = await fetch(`${BASE_URL}/api/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S3', 'x-forwarded-for': '10.0.4.5' },
        body: JSON.stringify({
            token: token2,
            rollNo: '25BCS003',
            deviceId: 'test-device-student-003-s2'
        })
    });
    if (!subRes2_3.ok) throw new Error('Failed to submit 25BCS003 in Session 2: ' + (await subRes2_3.text()));
    console.log('  ✓ 25BCS003 recorded present in Session 2.');

    // Stop Session 2
    await fetch(`${BASE_URL}/api/stop-session`, { method: 'POST' });

    // Download and verify both sessions in Excel
    const dlRes2 = await fetch(`${BASE_URL}/api/download`);
    const dlBuf2 = await dlRes2.arrayBuffer();
    const updatedWb2 = new ExcelJS.Workbook();
    await updatedWb2.xlsx.load(Buffer.from(dlBuf2));
    const wsCheck2 = updatedWb2.worksheets[0];

    console.log(`  ✓ Excel column count after Session 2: ${wsCheck2.columnCount}`);
    console.log(`  ✓ Col 2: '${wsCheck2.getRow(1).getCell(2).value}', Col 3: '${wsCheck2.getRow(1).getCell(3).value}'`);

    // In Col 2 (Session 1): 25BCS001=1, 25BCS002=1, 25BCS003=0
    // In Col 3 (Session 2): 25BCS001=0, 25BCS002=0, 25BCS003=1
    const s1_001 = wsCheck2.getRow(2).getCell(2).value;
    const s1_002 = wsCheck2.getRow(3).getCell(2).value;
    const s1_003 = wsCheck2.getRow(4).getCell(2).value;

    const s2_001 = wsCheck2.getRow(2).getCell(3).value;
    const s2_002 = wsCheck2.getRow(3).getCell(3).value;
    const s2_003 = wsCheck2.getRow(4).getCell(3).value;

    console.log(`  ✓ Session 1 values: 25BCS001=${s1_001}, 25BCS002=${s1_002}, 25BCS003=${s1_003}`);
    console.log(`  ✓ Session 2 values: 25BCS001=${s2_001}, 25BCS002=${s2_002}, 25BCS003=${s2_003}`);

    if (s1_001 !== 1 || s1_002 !== 1 || s1_003 !== 0 || s2_001 !== 0 || s2_002 !== 0 || s2_003 !== 1) {
        throw new Error('Values mismatch across multiple sessions! Overwrite detected.');
    }
    console.log('  ✓ VERIFIED: Previous session columns are completely preserved and never overwritten!');

    // 11. Multi-Class Independence Test
    console.log('\n[TEST 11] Testing Multi-Class Independence (Database_Management.xlsx)...');
    await fetch(`${BASE_URL}/api/classes/select`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId: dbmFileId })
    });
    const startResDbm = await fetch(`${BASE_URL}/api/reset`, { method: 'POST' });
    const startDataDbm = await startResDbm.json();
    console.log(`  ✓ Started session for Database_Management: '${startDataDbm.sessionColumn}'`);

    const tokenResDbm = await fetch(`${BASE_URL}/api/token`);
    const tokenDataDbm = await tokenResDbm.json();

    // Mark 25BCS010 in Database_Management
    await fetch(`${BASE_URL}/api/validate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S10', 'x-forwarded-for': '10.0.4.10' },
        body: JSON.stringify({ token: tokenDataDbm.token })
    });
    const subDbmRes = await fetch(`${BASE_URL}/api/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'User-Agent': CHROME_UA + ' S10', 'x-forwarded-for': '10.0.4.10' },
        body: JSON.stringify({
            token: tokenDataDbm.token,
            rollNo: '25BCS010',
            deviceId: 'test-device-dbm-010'
        })
    });
    if (!subDbmRes.ok) throw new Error('Failed to submit 25BCS010 in DBM: ' + (await subDbmRes.text()));
    await fetch(`${BASE_URL}/api/stop-session`, { method: 'POST' });

    // Download Database_Management and verify
    const dbmDlRes = await fetch(`${BASE_URL}/api/classes/${dbmFileId}/download`);
    const dbmDlBuf = await dbmDlRes.arrayBuffer();
    const dbmDlWb = new ExcelJS.Workbook();
    await dbmDlWb.xlsx.load(Buffer.from(dbmDlBuf));
    const dbmWsCheck = dbmDlWb.worksheets[0];

    const dbm_010 = dbmWsCheck.getRow(11).getCell(2).value;
    const dbm_001 = dbmWsCheck.getRow(2).getCell(2).value;
    console.log(`  ✓ Database Management records: 25BCS010=${dbm_010}, 25BCS001=${dbm_001}`);
    if (dbm_010 !== 1 || dbm_001 !== 0) {
        throw new Error('Database Management records mismatch!');
    }

    // Verify Data_Structures was NOT modified by Database_Management
    const dsVerifyRes = await fetch(`${BASE_URL}/api/classes/${dsFileId}/download`);
    const dsVerifyBuf = await dsVerifyRes.arrayBuffer();
    const dsVerifyWb = new ExcelJS.Workbook();
    await dsVerifyWb.xlsx.load(Buffer.from(dsVerifyBuf));
    const dsWsVerify = dsVerifyWb.worksheets[0];
    if (dsWsVerify.columnCount !== 3) {
        throw new Error(`Data_Structures column count corrupted: expected 3, got ${dsWsVerify.columnCount}`);
    }
    console.log('  ✓ VERIFIED: Multiple classes maintain completely independent attendance histories!');

    // 12. SECTION 11 TESTS: Interactive Spreadsheet Viewer & Manual Attendance Editing
    console.log('\n[TEST 12] Testing Section 11: Interactive Spreadsheet Data API (GET /api/classes/:fileId/sheet)...');
    const sheetRes = await fetch(`${BASE_URL}/api/classes/${dsFileId}/sheet`);
    const sheetData = await sheetRes.json();
    if (!sheetRes.ok || !sheetData.success) {
        throw new Error('Failed to fetch spreadsheet data: ' + sheetData.error);
    }
    console.log(`  ✓ Spreadsheet data received: ${sheetData.columns.length} columns, ${sheetData.rows.length} rows.`);
    if (sheetData.rows.length !== 215) {
        throw new Error(`Expected 215 student rows in spreadsheet, got ${sheetData.rows.length}`);
    }

    // 12b. Manual Editing: Change 25BCS001 from 1 to 0, and change 25BCS005 from 0 to 1
    console.log('\n[TEST 12b] Testing Manual Attendance Editing & Persistence (PUT /api/classes/:fileId/sheet)...');
    const manualEditRes = await fetch(`${BASE_URL}/api/classes/${dsFileId}/sheet`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            updates: [
                { rollNo: '25BCS001', colNumber: 2, value: 0 }, // Changed 1 -> 0
                { rollNo: '25BCS005', colNumber: 2, value: 1 }  // Changed 0 -> 1
            ]
        })
    });
    const manualEditData = await manualEditRes.json();
    if (!manualEditRes.ok || !manualEditData.success) {
        throw new Error('Failed to save manual edits: ' + manualEditData.error);
    }
    console.log('  ✓ Manual edits saved to disk and database:', manualEditData.message);

    // Verify disk Excel file contains manual corrections
    const verifyManualDl = await fetch(`${BASE_URL}/api/classes/${dsFileId}/download`);
    const verifyManualBuf = await verifyManualDl.arrayBuffer();
    const verifyManualWb = new ExcelJS.Workbook();
    await verifyManualWb.xlsx.load(Buffer.from(verifyManualBuf));
    const vmWs = verifyManualWb.worksheets[0];

    const edited_001 = vmWs.getRow(2).getCell(2).value; // 25BCS001 at col 2
    const edited_005 = vmWs.getRow(6).getCell(2).value; // 25BCS005 at col 2
    console.log(`  ✓ Verified manual edits in Excel file: 25BCS001=${edited_001} (was 1), 25BCS005=${edited_005} (was 0)`);
    if (edited_001 !== 0 || edited_005 !== 1) {
        throw new Error(`Manual edits were not persisted in Excel! 001=${edited_001}, 005=${edited_005}`);
    }

    // 12c. Requirement 11.D: Upcoming Attendance must use the latest edited spreadsheet
    console.log('\n[TEST 12c] Testing Upcoming Attendance with Manually Edited Roster (Requirement 11.D)...');
    // Select Data_Structures again
    await fetch(`${BASE_URL}/api/classes/select`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId: dsFileId })
    });

    // Start Session 3
    const startRes3 = await fetch(`${BASE_URL}/api/reset`, { method: 'POST' });
    const startData3 = await startRes3.json();
    if (!startRes3.ok || !startData3.success) throw new Error('Failed to start session 3: ' + startData3.error);
    console.log(`  ✓ Upcoming Session 3 started: '${startData3.sessionColumn}'`);

    // Verify that the new column was added WITHOUT reverting manual corrections in Session 1
    const verifyUpcomingDl = await fetch(`${BASE_URL}/api/classes/${dsFileId}/download`);
    const verifyUpcomingBuf = await verifyUpcomingDl.arrayBuffer();
    const verifyUpcomingWb = new ExcelJS.Workbook();
    await verifyUpcomingWb.xlsx.load(Buffer.from(verifyUpcomingBuf));
    const vuWs = verifyUpcomingWb.worksheets[0];

    console.log(`  ✓ Total columns in Data_Structures now: ${vuWs.columnCount}`);
    const hist_001 = vuWs.getRow(2).getCell(2).value;
    const hist_005 = vuWs.getRow(6).getCell(2).value;
    const new_001 = vuWs.getRow(2).getCell(4).value; // Col 4 is Session 3
    const new_005 = vuWs.getRow(6).getCell(4).value;

    console.log(`  ✓ Preserved Historical Session 1: 25BCS001=${hist_001}, 25BCS005=${hist_005}`);
    console.log(`  ✓ Initialized New Session 3: 25BCS001=${new_001}, 25BCS005=${new_005}`);

    if (hist_001 !== 0 || hist_005 !== 1) {
        throw new Error('Upcoming session reverted historical manual corrections!');
    }
    if (new_001 !== 0 || new_005 !== 0) {
        throw new Error('New session did not initialize students to 0!');
    }
    console.log('  ✓ VERIFIED: Upcoming session seamlessly builds upon the latest edited spreadsheet without reverting historical records!');

    await fetch(`${BASE_URL}/api/stop-session`, { method: 'POST' });

    console.log('\n====================================================');
    console.log(' ALL TESTS PASSED SUCCESSFULLY! 100% SPEC COMPLIANT');
    console.log('====================================================\n');
}

runTests().catch(err => {
    console.error('\n❌ TEST FAILED:', err);
    process.exit(1);
});
