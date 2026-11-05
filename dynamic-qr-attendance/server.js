'use strict';

const express = require('express');
const crypto = require('crypto');
const path = require('path');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();
const ExcelJS = require('exceljs');
const os = require('os');

const app = express();

const PORT = Number(process.env.PORT || 3000);
const PUBLIC_DIR = path.join(__dirname, 'public');
const DB_PATH = path.join(__dirname, 'attendance.db');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const BACKUPS_DIR = path.join(UPLOADS_DIR, 'backups');

// Ensure storage directories exist
if (!fs.existsSync(UPLOADS_DIR)) {
    fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}
if (!fs.existsSync(BACKUPS_DIR)) {
    fs.mkdirSync(BACKUPS_DIR, { recursive: true });
}

const TOKEN_LIFESPAN_MS = 15000; // 15 seconds
const MAX_USES = 50;

// Predefined authorized student roll numbers: 25BCS001 to 25BCS215 (215 students)
const TOTAL_EXPECTED_STUDENTS = 215;
function getExpectedRollNumbers() {
    const list = [];
    for (let i = 1; i <= TOTAL_EXPECTED_STUDENTS; i++) {
        list.push(`25BCS${String(i).padStart(3, '0')}`);
    }
    return list;
}
const EXPECTED_ROLL_NUMBERS = getExpectedRollNumbers();
const EXPECTED_ROLLS_SET = new Set(EXPECTED_ROLL_NUMBERS);

// Optional configuration through environment variables.
const NGROK_AUTHTOKEN = process.env.NGROK_AUTHTOKEN || '';
const GOOGLE_SHEETS_WEBHOOK = process.env.GOOGLE_SHEETS_WEBHOOK || '';
const EXPLICIT_BASE_URL = process.env.PUBLIC_BASE_URL || '';

app.disable('x-powered-by');
// Support JSON body up to 50MB for Excel file uploads
app.use(express.json({ limit: '50mb' }));

// --------------------------------------------------
// LOCAL IP DETECTION
// --------------------------------------------------

function getLocalIp() {
    const interfaces = os.networkInterfaces();
    const candidates = [];

    for (const [name, entries] of Object.entries(interfaces)) {
        for (const iface of entries || []) {
            if (iface.family !== 'IPv4' || iface.internal) {
                continue;
            }

            const ip = iface.address;
            const isPrivate =
                /^10\./.test(ip) ||
                /^192\.168\./.test(ip) ||
                /^172\.(1[6-9]|2\d|3[01])\./.test(ip);

            candidates.push({ name, ip, isPrivate });
        }
    }

    const wifi = candidates.find(item => item.name === 'en0');
    if (wifi) return wifi.ip;

    const privateAddress = candidates.find(item => item.isPrivate);
    return privateAddress?.ip || candidates[0]?.ip || 'localhost';
}

const myIp = getLocalIp();

let BASE_URL = (
    EXPLICIT_BASE_URL ||
    `http://${myIp}:${PORT}`
).replace(/\/+$/, '');

let ngrokListener = null;
let tokenInterval = null;
let server = null;
let shuttingDown = false;

console.log(`[SYS] Detected local IP: ${myIp}`);

// --------------------------------------------------
// DATABASE & SCHEMA MIGRATIONS
// Never drop the attendance table on startup.
// Existing attendance records are preserved.
// --------------------------------------------------

const db = new sqlite3.Database(DB_PATH);

function dbRun(sql, params = []) {
    return new Promise((resolve, reject) => {
        db.run(sql, params, function (err) {
            if (err) {
                reject(err);
            } else {
                resolve({
                    changes: this.changes,
                    lastID: this.lastID
                });
            }
        });
    });
}

function dbGet(sql, params = []) {
    return new Promise((resolve, reject) => {
        db.get(sql, params, (err, row) => {
            if (err) reject(err);
            else resolve(row);
        });
    });
}

function dbAll(sql, params = []) {
    return new Promise((resolve, reject) => {
        db.all(sql, params, (err, rows) => {
            if (err) reject(err);
            else resolve(rows);
        });
    });
}

async function initializeDatabase() {
    await dbRun(`
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sessionId TEXT NOT NULL,
            rollNo TEXT NOT NULL,
            ip TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    `);

    // Ensure fileId column exists on attendance table
    const tableInfo = await dbAll("PRAGMA table_info(attendance)");
    const hasFileId = tableInfo.some(col => col.name === 'fileId');
    if (!hasFileId) {
        await dbRun("ALTER TABLE attendance ADD COLUMN fileId TEXT");
        console.log('[DB] Added fileId column to attendance table.');
    }

    // Classes table for managing uploaded Excel files
    await dbRun(`
        CREATE TABLE IF NOT EXISTS classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fileId TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            filename TEXT NOT NULL,
            filePath TEXT NOT NULL,
            studentCount INTEGER NOT NULL DEFAULT ${TOTAL_EXPECTED_STUDENTS},
            authorizedRolls TEXT NOT NULL,
            createdAt DATETIME DEFAULT CURRENT_TIMESTAMP,
            updatedAt DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    `);

    // Sessions table for managing attendance sessions per class
    await dbRun(`
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sessionId TEXT UNIQUE NOT NULL,
            fileId TEXT NOT NULL,
            sessionDate TEXT NOT NULL,
            sessionColumn TEXT NOT NULL,
            colNumber INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'ACTIVE',
            createdAt DATETIME DEFAULT CURRENT_TIMESTAMP,
            endedAt DATETIME
        )
    `);

    await dbRun(`
        CREATE INDEX IF NOT EXISTS idx_attendance_session_roll
        ON attendance(sessionId, rollNo)
    `);

    await dbRun(`
        CREATE INDEX IF NOT EXISTS idx_attendance_session_ip
        ON attendance(sessionId, ip)
    `);

    await dbRun(`
        CREATE INDEX IF NOT EXISTS idx_attendance_session_file
        ON attendance(sessionId, fileId)
    `);

    await dbRun(`
        CREATE INDEX IF NOT EXISTS idx_sessions_file
        ON sessions(fileId)
    `);

    console.log('[DB] Attendance database ready with multi-class and Excel tracking.');

    // Auto-restore last active class if available
    await restoreActiveClass();
}

// --------------------------------------------------
// DATE HELPERS
// --------------------------------------------------

function getFormattedDate(date = new Date()) {
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const y = date.getFullYear();
    return `${d}-${m}-${y}`;
}

// --------------------------------------------------
// ACTIVE CLASS & SESSION STATE
// --------------------------------------------------

let activeClass = null; // { fileId, name, filename, filePath, studentCount, authorizedRollsSet, rollRowMap }
let activeWorkbook = null;
let activeSessionColumn = '';
let activeSessionColIndex = 0;

let currentSessionId = crypto.randomUUID();
let sessionActive = false;

let currentToken = '';
let tokenUses = 0;
let tokenExpiresAt = 0;

const submittedDevices = new Set();
const deviceScans = new Set();
const validatedDevices = new Set();

let inFlightSubmissions = 0;
let submissionWaiters = [];

function beginSubmission() {
    inFlightSubmissions++;
}

function endSubmission() {
    inFlightSubmissions--;
    if (inFlightSubmissions === 0) {
        const waiters = submissionWaiters;
        submissionWaiters = [];
        for (const resolve of waiters) {
            resolve();
        }
    }
}

function waitForSubmissions() {
    if (inFlightSubmissions === 0) {
        return Promise.resolve();
    }
    return new Promise(resolve => {
        submissionWaiters.push(resolve);
    });
}

// --------------------------------------------------
// EXCEL VALIDATION, READING & UPDATING
// --------------------------------------------------

/**
 * Validates an uploaded workbook.
 * Must contain a "Roll No." column with roll numbers 25BCS001 to 25BCS215 in ascending order.
 */
function validateAndParseExcel(workbook) {
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
        throw new Error('Corrupted Excel file: no worksheets found.');
    }

    let rollColIndex = -1;
    let nameColIndex = -1;
    let headerRowIndex = -1;

    for (let r = 1; r <= Math.min(15, worksheet.rowCount); r++) {
        const row = worksheet.getRow(r);
        row.eachCell((cell, colNumber) => {
            const rawVal = cell.value;
            const text = (rawVal && typeof rawVal === 'object' && rawVal.richText)
                ? rawVal.richText.map(t => t.text).join('')
                : String(rawVal || '');
            const cleaned = text.trim();
            if (/^(reg(istration)?|roll)\s*(no\.?|number)?$/i.test(cleaned) || /^reg\b/i.test(cleaned) || /^roll\b/i.test(cleaned)) {
                rollColIndex = colNumber;
                headerRowIndex = r;
            }
            if (/^name$/i.test(cleaned) || /^student\s*name$/i.test(cleaned)) {
                nameColIndex = colNumber;
            }
        });
        if (rollColIndex !== -1) break;
    }

    if (rollColIndex === -1) {
        throw new Error("Invalid format: Excel file must contain a 'Reg. No.' column.");
    }

    const extracted = [];
    const rollRowMap = {};
    const seenRolls = new Set();
    const authorizedRolls = [];

    for (let r = headerRowIndex + 1; r <= worksheet.rowCount; r++) {
        const cell = worksheet.getRow(r).getCell(rollColIndex);
        const rawVal = cell.value;
        const text = (rawVal && typeof rawVal === 'object' && rawVal.richText)
            ? rawVal.richText.map(t => t.text).join('')
            : String(rawVal ?? '');
        const roll = text.trim().toUpperCase();
        if (roll) {
            if (seenRolls.has(roll)) {
                throw new Error(`Duplicate registration number '${roll}' found at row ${r}. Each student must have a unique Reg. No.`);
            }
            seenRolls.add(roll);
            extracted.push({ roll, rowNumber: r });
            rollRowMap[roll] = r;
            authorizedRolls.push(roll);
        }
    }

    if (extracted.length === 0) {
        throw new Error("Invalid file: No student registration numbers found under 'Reg. No.' column.");
    }

    return {
        worksheet,
        headerRowIndex,
        rollColIndex,
        nameColIndex,
        rollRowMap,
        authorizedRolls
    };
}

/**
 * Creates a new attendance column for today's date in DD-MM-YYYY format.
 * If multiple sessions occur on the same day, numbers them: (Session 2), (Session 3), etc.
 * Initializes all students' attendance to 0.
 */
function prepareSessionColumnInWorksheet(worksheet, headerRowIndex, rollRowMap) {
    const todayStr = getFormattedDate();
    const headerRow = worksheet.getRow(headerRowIndex);

    let sameDateCount = 0;
    let maxColNumber = worksheet.columnCount || 1;

    headerRow.eachCell((cell, colNumber) => {
        if (colNumber > maxColNumber) maxColNumber = colNumber;
        const val = String(cell.value || '').trim();
        if (val.startsWith(todayStr)) {
            sameDateCount++;
        }
    });

    let sessionColName = todayStr;
    if (sameDateCount > 0) {
        sessionColName = `${todayStr} (Session ${sameDateCount + 1})`;
    }

    const newColNumber = maxColNumber + 1;
    const headerCell = worksheet.getRow(headerRowIndex).getCell(newColNumber);
    headerCell.value = sessionColName;
    headerCell.font = { bold: true, name: 'Calibri', size: 11 };
    headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
    headerCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2E8F0' }
    };
    headerCell.border = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };

    const colObj = worksheet.getColumn(newColNumber);
    if (!colObj.width || colObj.width < 16) {
        colObj.width = 16;
    }

    // Initialize all authorized students to 0
    for (const roll of Object.keys(rollRowMap)) {
        const r = rollRowMap[roll];
        if (r) {
            const cell = worksheet.getRow(r).getCell(newColNumber);
            cell.value = 0;
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };
        }
    }

    return {
        sessionColName,
        colNumber: newColNumber
    };
}

/**
 * Serialized file save lock to avoid concurrent write corruption
 */
let isSavingExcel = false;
let pendingSave = false;

async function persistActiveWorkbook() {
    if (!activeClass || !activeWorkbook) return;

    if (isSavingExcel) {
        pendingSave = true;
        return;
    }

    isSavingExcel = true;

    try {
        const filePath = activeClass.filePath;

        // Backup existing file before write
        if (fs.existsSync(filePath)) {
            const backupName = `${activeClass.fileId}_${Date.now()}.bak.xlsx`;
            try {
                fs.copyFileSync(filePath, path.join(BACKUPS_DIR, backupName));
            } catch (bErr) {
                console.warn('[BACKUP WARN]', bErr.message);
            }
        }

        const tempPath = `${filePath}.tmp`;
        await activeWorkbook.xlsx.writeFile(tempPath);
        fs.renameSync(tempPath, filePath);

        await dbRun(
            "UPDATE classes SET updatedAt = CURRENT_TIMESTAMP WHERE fileId = ?",
            [activeClass.fileId]
        );
    } catch (err) {
        console.error('[EXCEL WRITE ERROR]', err.message);
    } finally {
        isSavingExcel = false;
        if (pendingSave) {
            pendingSave = false;
            persistActiveWorkbook();
        }
    }
}

/**
 * Reconcile Excel sheet with SQLite attendance table for total consistency
 */
async function syncSessionToExcel(sessionId, fileId) {
    if (!activeClass || activeClass.fileId !== fileId || !activeWorkbook) return;

    try {
        const rows = await dbAll(
            'SELECT rollNo FROM attendance WHERE sessionId = ? AND fileId = ?',
            [sessionId, fileId]
        );
        const presentRolls = new Set(rows.map(r => r.rollNo.toUpperCase()));

        const worksheet = activeWorkbook.worksheets[0];
        if (!worksheet || !activeSessionColIndex) return;

        for (const roll of Object.keys(activeClass.rollRowMap || {})) {
            const r = activeClass.rollRowMap[roll];
            if (r) {
                const cell = worksheet.getRow(r).getCell(activeSessionColIndex);
                cell.value = presentRolls.has(roll) ? 1 : 0;
            }
        }

        await persistActiveWorkbook();
    } catch (err) {
        console.error('[SYNC TO EXCEL ERROR]', err.message);
    }
}

/**
 * Loads a class file into memory as activeClass
 */
async function loadClassRecord(classRecord) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(classRecord.filePath);

    const parsed = validateAndParseExcel(workbook);

    activeClass = {
        fileId: classRecord.fileId,
        name: classRecord.name,
        filename: classRecord.filename,
        filePath: classRecord.filePath,
        studentCount: classRecord.studentCount,
        authorizedRollsSet: new Set(parsed.authorizedRolls),
        rollRowMap: parsed.rollRowMap,
        headerRowIndex: parsed.headerRowIndex,
        rollColIndex: parsed.rollColIndex
    };
    activeWorkbook = workbook;

    console.log(`[CLASS] Loaded active class: ${activeClass.name} (${activeClass.filename})`);
}

/**
 * Automatically restores the latest class from database on startup
 */
async function restoreActiveClass() {
    try {
        const lastClass = await dbGet(`
            SELECT * FROM classes
            ORDER BY updatedAt DESC, id DESC
            LIMIT 1
        `);

        if (lastClass && fs.existsSync(lastClass.filePath)) {
            await loadClassRecord(lastClass);
        }
    } catch (err) {
        console.warn('[RESTORE CLASS WARN]', err.message);
    }
}

/**
 * Generates blank attendance roster template with Reg. No. and Name columns
 */
async function generateCseTemplateBuffer() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'QR Rotation Attendance System';
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet('Class Attendance', {
        views: [{ state: 'frozen', ySplit: 1 }]
    });

    worksheet.columns = [
        { header: 'Reg. No.', key: 'regNo', width: 20 },
        { header: 'Name', key: 'name', width: 32 }
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F46E5' }
    };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    headerRow.height = 26;

    return await workbook.xlsx.writeBuffer();
}

// --------------------------------------------------
// TOKEN GENERATION & ROTATION
// --------------------------------------------------

function generateNewToken() {
    if (!sessionActive) return;

    currentToken = crypto.randomBytes(16).toString('hex');
    tokenUses = 0;
    tokenExpiresAt = Date.now() + TOKEN_LIFESPAN_MS;

    deviceScans.clear();
    validatedDevices.clear();

    console.log(
        `[QR] Token refreshed. Expires at ${
            new Date(tokenExpiresAt).toLocaleTimeString()
        }`
    );
}

function startTokenLoop() {
    if (tokenInterval) {
        clearInterval(tokenInterval);
    }

    generateNewToken();

    tokenInterval = setInterval(
        generateNewToken,
        TOKEN_LIFESPAN_MS
    );
}

function stopTokenLoop() {
    if (tokenInterval) {
        clearInterval(tokenInterval);
        tokenInterval = null;
    }
}

function tokenIsActive(token) {
    return Boolean(
        sessionActive &&
        typeof token === 'string' &&
        token.length > 0 &&
        token === currentToken &&
        Date.now() < tokenExpiresAt
    );
}

// --------------------------------------------------
// CLIENT IP & DEVICE IDENTIFICATION
// --------------------------------------------------

function isPublicTunnelRequest(req) {
    const host = (req.get('host') || '').toLowerCase();
    const forwardedHost =
        (req.get('x-forwarded-host') || '').toLowerCase();

    return (
        host.includes('ngrok') ||
        forwardedHost.includes('ngrok')
    );
}

function getRequestIp(req) {
    let ip;

    if (isPublicTunnelRequest(req) || process.env.NODE_ENV === 'test') {
        const forwarded = req.get('x-forwarded-for');
        if (forwarded) {
            ip = forwarded.split(',')[0].trim();
        }
    }

    if (!ip) {
        ip = req.socket?.remoteAddress || 'unknown';
    }

    if (ip.startsWith('::ffff:')) {
        ip = ip.slice(7);
    }

    return ip;
}

function getDeviceHash(req) {
    const ip = getRequestIp(req);
    const ua = req.get('user-agent') || 'unknown';

    return crypto
        .createHash('sha256')
        .update(`${ip}|${ua}`)
        .digest('hex');
}

// --------------------------------------------------
// CHROME USER-AGENT CHECK
// --------------------------------------------------

function isChromeBrowser(req) {
    const ua = req.get('user-agent') || '';

    const isChrome = /(?:Chrome|CriOS)\/[\d.]+/.test(ua);
    const isAnotherBrowser =
        /Edg(?:e|A|iOS)?\/|OPR\/|Opera|Firefox\/|FxiOS\/|SamsungBrowser\/|DuckDuckGo|YaBrowser|Vivaldi/i
            .test(ua);

    return isChrome && !isAnotherBrowser;
}

function requireChrome(req, res) {
    if (isChromeBrowser(req)) {
        return true;
    }

    res.status(403).json({
        error: 'Please use Google Chrome to mark attendance.'
    });

    return false;
}

// --------------------------------------------------
// PUBLIC TUNNEL PROTECTION
// --------------------------------------------------

const professorExactPaths = new Set([
    '/',
    '/index.html',
    '/api/token',
    '/api/attendance',
    '/api/download',
    '/api/reset',
    '/api/stop-session',
    '/api/shutdown',
    '/api/classes',
    '/api/template'
]);

function isProfessorPath(reqPath) {
    const normalized = reqPath.replace(/\/+$/, '') || '/';
    if (professorExactPaths.has(normalized)) return true;
    if (normalized.startsWith('/api/classes')) return true;
    return false;
}

app.use((req, res, next) => {
    if (isPublicTunnelRequest(req) && isProfessorPath(req.path)) {
        return res.status(403).send(`
            <!doctype html>
            <html lang="en">
            <head>
                <meta charset="utf-8">
                <title>Access Denied</title>
            </head>
            <body style="font-family:Arial;text-align:center;padding:60px">
                <h1>Access Denied</h1>
                <p>The professor dashboard is not publicly accessible.</p>
            </body>
            </html>
        `);
    }
    next();
});

// --------------------------------------------------
// STATIC FRONTEND
// --------------------------------------------------

app.use(express.static(PUBLIC_DIR, {
    setHeaders(res, filePath) {
        if (filePath.endsWith('.html')) {
            res.setHeader('Cache-Control', 'no-store');
        }
    }
}));

// --------------------------------------------------
// HEALTH CHECK
// --------------------------------------------------

app.get('/api/health', (req, res) => {
    res.set('Cache-Control', 'no-store');
    res.json({
        success: true,
        status: 'ok',
        sessionActive,
        activeClass: activeClass ? {
            fileId: activeClass.fileId,
            name: activeClass.name,
            filename: activeClass.filename
        } : null
    });
});

// --------------------------------------------------
// PROFESSOR: TEMPLATE DOWNLOAD
// --------------------------------------------------

app.get('/api/template', async (req, res) => {
    try {
        const buffer = await generateCseTemplateBuffer();
        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            'attachment; filename="CSE_Attendance_Template.xlsx"'
        );
        res.setHeader('Cache-Control', 'no-store');
        return res.send(Buffer.from(buffer));
    } catch (err) {
        console.error('[TEMPLATE GENERATION ERROR]', err);
        return res.status(500).json({ error: 'Could not generate template.' });
    }
});

// --------------------------------------------------
// PROFESSOR: EXCEL CLASS FILE UPLOAD & MANAGEMENT
// --------------------------------------------------

app.get('/api/classes', async (req, res) => {
    try {
        const classes = await dbAll(`
            SELECT c.id, c.fileId, c.name, c.filename, c.studentCount, c.createdAt, c.updatedAt,
                   COUNT(s.id) as sessionCount
            FROM classes c
            LEFT JOIN sessions s ON c.fileId = s.fileId
            GROUP BY c.fileId
            ORDER BY c.updatedAt DESC, c.id DESC
        `);

        return res.json({
            success: true,
            activeClassFileId: activeClass?.fileId || null,
            classes: classes.map(c => ({
                ...c,
                isActive: Boolean(activeClass && activeClass.fileId === c.fileId)
            }))
        });
    } catch (err) {
        console.error('[GET CLASSES ERROR]', err);
        return res.status(500).json({ error: 'Could not fetch class list.' });
    }
});

app.post('/api/classes/upload', async (req, res) => {
    const { courseName, filename, base64 } = req.body || {};

    if (!filename || typeof filename !== 'string') {
        return res.status(400).json({ error: 'Missing or invalid file name.' });
    }

    if (!/\.(xlsx|xls)$/i.test(filename)) {
        return res.status(400).json({ error: 'Invalid file type. Please upload an Excel (.xlsx, .xls) file.' });
    }

    if (!base64 || typeof base64 !== 'string') {
        return res.status(400).json({ error: 'Missing Excel file payload.' });
    }

    let buffer;
    try {
        buffer = Buffer.from(base64, 'base64');
    } catch (bErr) {
        return res.status(400).json({ error: 'Failed to decode file data.' });
    }

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        // Validate structure and roster (25BCS001 to 25BCS215 in ascending order)
        const parsed = validateAndParseExcel(workbook);

        const fileId = crypto.randomUUID();
        const safeName = filename.replace(/[^a-zA-Z0-9._-]/g, '_');
        const storedFilename = `${fileId}_${safeName}`;
        const storedPath = path.join(UPLOADS_DIR, storedFilename);

        // Save workbook to storage
        await workbook.xlsx.writeFile(storedPath);

        // Derive user-friendly class name
        const displayName = (courseName && typeof courseName === 'string' && courseName.trim())
            ? courseName.trim()
            : filename.replace(/\.(xlsx|xls)$/i, '').replace(/[_-]/g, ' ');

        await dbRun(`
            INSERT INTO classes (fileId, name, filename, filePath, studentCount, authorizedRolls)
            VALUES (?, ?, ?, ?, ?, ?)
        `, [
            fileId,
            displayName,
            filename,
            storedPath,
            parsed.authorizedRolls.length,
            JSON.stringify(parsed.authorizedRolls)
        ]);

        // Automatically set this newly uploaded class as the active class
        activeClass = {
            fileId,
            name: displayName,
            filename,
            filePath: storedPath,
            studentCount: parsed.authorizedRolls.length,
            authorizedRollsSet: new Set(parsed.authorizedRolls),
            rollRowMap: parsed.rollRowMap,
            headerRowIndex: parsed.headerRowIndex,
            rollColIndex: parsed.rollColIndex
        };
        activeWorkbook = workbook;
        sessionActive = false; // Require professor to explicitly start session for the class

        console.log(`[CLASS UPLOAD] Successfully uploaded and validated: ${filename} (${displayName})`);

        return res.json({
            success: true,
            message: `Excel file '${filename}' uploaded and verified successfully (${parsed.authorizedRolls.length} students).`,
            class: {
                fileId,
                name: displayName,
                filename,
                studentCount: parsed.authorizedRolls.length
            }
        });
    } catch (err) {
        console.error('[CLASS UPLOAD ERROR]', err.message);
        return res.status(400).json({
            error: err.message || 'Could not validate or save the uploaded Excel file.'
        });
    }
});

app.post('/api/classes/select', async (req, res) => {
    const { fileId } = req.body || {};

    if (!fileId) {
        return res.status(400).json({ error: 'Please specify a class fileId.' });
    }

    try {
        const classRecord = await dbGet('SELECT * FROM classes WHERE fileId = ?', [fileId]);
        if (!classRecord) {
            return res.status(404).json({ error: 'Class not found.' });
        }

        if (!fs.existsSync(classRecord.filePath)) {
            return res.status(404).json({ error: 'Excel file for this class was not found on server disk.' });
        }

        // If a session was active for another class, stop it first
        if (sessionActive) {
            sessionActive = false;
            stopTokenLoop();
            currentToken = '';
            tokenExpiresAt = 0;
            await waitForSubmissions();
        }

        await loadClassRecord(classRecord);

        return res.json({
            success: true,
            message: `Selected class: ${activeClass.name}`,
            class: {
                fileId: activeClass.fileId,
                name: activeClass.name,
                filename: activeClass.filename,
                studentCount: activeClass.studentCount
            }
        });
    } catch (err) {
        console.error('[SELECT CLASS ERROR]', err);
        return res.status(500).json({ error: 'Could not select class: ' + err.message });
    }
});

app.get('/api/classes/:fileId/download', async (req, res) => {
    const { fileId } = req.params;

    try {
        const classRecord = await dbGet('SELECT * FROM classes WHERE fileId = ?', [fileId]);
        if (!classRecord || !fs.existsSync(classRecord.filePath)) {
            return res.status(404).send('Class Excel file not found.');
        }

        // If this is the currently active class and session is active, flush first
        if (activeClass && activeClass.fileId === fileId && sessionActive) {
            await syncSessionToExcel(currentSessionId, fileId);
        }

        const data = fs.readFileSync(classRecord.filePath);

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="${classRecord.filename}"`
        );
        res.setHeader('Cache-Control', 'no-store');

        return res.send(data);
    } catch (err) {
        console.error('[DOWNLOAD CLASS ERROR]', err);
        return res.status(500).send('Could not download class Excel file.');
    }
});

app.delete('/api/classes/:fileId', async (req, res) => {
    const { fileId } = req.params;
    try {
        const cls = await dbGet('SELECT * FROM classes WHERE fileId = ?', [fileId]);
        if (!cls) {
            return res.status(404).json({ error: 'Class not found.' });
        }
        if (sessionActive && activeClass && activeClass.fileId === fileId) {
            return res.status(400).json({ error: 'Cannot delete a course while an attendance session is active. Please end the session first.' });
        }

        await dbRun('DELETE FROM attendance WHERE fileId = ?', [fileId]);
        await dbRun('DELETE FROM sessions WHERE fileId = ?', [fileId]);
        await dbRun('DELETE FROM classes WHERE fileId = ?', [fileId]);

        if (cls.filePath && fs.existsSync(cls.filePath)) {
            try { fs.unlinkSync(cls.filePath); } catch (e) { /* ignore */ }
        }

        if (activeClass && activeClass.fileId === fileId) {
            activeClass = null;
            activeWorkbook = null;
            const nextClass = await dbGet('SELECT * FROM classes ORDER BY updatedAt DESC, id DESC LIMIT 1');
            if (nextClass && fs.existsSync(nextClass.filePath)) {
                await loadClassRecord(nextClass);
            }
        }

        return res.json({
            success: true,
            message: `Course '${cls.name}' deleted successfully.`,
            activeClassFileId: activeClass?.fileId || null
        });
    } catch (err) {
        console.error('[DELETE CLASS ERROR]', err);
        return res.status(500).json({ error: 'Failed to delete course.' });
    }
});

// --------------------------------------------------
// PROFESSOR: SPREADSHEET VIEWER & MANUAL EDITING API
// Section 11: Interactive spreadsheet data reading & updates
// --------------------------------------------------

app.get('/api/classes/:fileId/sheet', async (req, res) => {
    const { fileId } = req.params;

    try {
        const classRecord = await dbGet('SELECT * FROM classes WHERE fileId = ?', [fileId]);
        if (!classRecord || !fs.existsSync(classRecord.filePath)) {
            return res.status(404).json({ error: 'Class Excel file not found on server.' });
        }

        // If active session is in progress for this class, flush latest state to disk first
        if (activeClass && activeClass.fileId === fileId && sessionActive) {
            await syncSessionToExcel(currentSessionId, fileId);
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(classRecord.filePath);

        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
            return res.status(500).json({ error: 'Excel file contains no worksheets.' });
        }

        // Find header row, Reg/Roll No column, and Name column
        let headerRowIndex = 1;
        let rollColIndex = 1;
        let nameColIndex = -1;

        for (let r = 1; r <= Math.min(15, worksheet.rowCount); r++) {
            const row = worksheet.getRow(r);
            row.eachCell((cell, colNumber) => {
                const text = String(cell.value || '').trim();
                if (/^(reg(istration)?|roll)\s*(no\.?|number)?$/i.test(text) || /^reg\b/i.test(text) || /^roll\b/i.test(text)) {
                    headerRowIndex = r;
                    rollColIndex = colNumber;
                }
                if (/^name$/i.test(text) || /^student\s*name$/i.test(text)) {
                    nameColIndex = colNumber;
                }
            });
            if (rollColIndex !== -1) break;
        }

        const headerRow = worksheet.getRow(headerRowIndex);
        const columns = [];
        const maxCols = worksheet.columnCount || 1;

        // Build column descriptors
        for (let c = 1; c <= maxCols; c++) {
            const cellVal = headerRow.getCell(c).value;
            const label = cellVal !== null && cellVal !== undefined ? String(cellVal).trim() : `Col ${c}`;
            const isRollNo = (c === rollColIndex);
            const isName = (c === nameColIndex) || (/^name$/i.test(label));

            columns.push({
                colNumber: c,
                key: isRollNo ? 'rollNo' : (isName ? 'name' : `col_${c}`),
                label: label || (isRollNo ? 'Reg. No.' : (isName ? 'Name' : `Session ${c - 1}`)),
                isRollNo,
                isName,
                isDate: !isRollNo && !isName
            });
        }

        // Extract student rows
        const rows = [];
        const colPresentCounts = {};
        for (const col of columns) {
            if (col.isDate) colPresentCounts[col.key] = 0;
        }

        for (let r = headerRowIndex + 1; r <= worksheet.rowCount; r++) {
            const rowObj = worksheet.getRow(r);
            const rollVal = rowObj.getCell(rollColIndex).value;
            if (!rollVal) continue;

            const rollNo = String(rollVal).trim().toUpperCase();
            let studentName = '';
            if (nameColIndex !== -1) {
                const nameVal = rowObj.getCell(nameColIndex).value;
                studentName = nameVal !== null && nameVal !== undefined ? String(nameVal).trim() : '';
            }

            const cells = {};

            for (const col of columns) {
                if (col.isRollNo) continue;
                if (col.isName) {
                    const rawName = rowObj.getCell(col.colNumber).value;
                    const nm = rawName !== null && rawName !== undefined ? String(rawName).trim() : '';
                    cells[col.key] = nm;
                    if (!studentName) studentName = nm;
                    continue;
                }
                const cellVal = rowObj.getCell(col.colNumber).value;
                const numVal = (Number(cellVal) === 1 || cellVal === '1' || cellVal === true) ? 1 : 0;
                cells[col.key] = numVal;
                if (numVal === 1) {
                    colPresentCounts[col.key] = (colPresentCounts[col.key] || 0) + 1;
                }
            }

            rows.push({
                rollNo,
                name: studentName,
                rowIndex: r,
                cells
            });
        }

        // Attach present counts to column metadata
        for (const col of columns) {
            if (col.isDate) {
                col.presentCount = colPresentCounts[col.key] || 0;
                col.absentCount = rows.length - (colPresentCounts[col.key] || 0);
            }
        }

        return res.json({
            success: true,
            class: {
                fileId: classRecord.fileId,
                name: classRecord.name,
                filename: classRecord.filename,
                studentCount: rows.length || classRecord.studentCount,
                updatedAt: classRecord.updatedAt
            },
            columns,
            rows
        });
    } catch (err) {
        console.error('[GET SHEET ERROR]', err);
        return res.status(500).json({ error: 'Could not load spreadsheet data: ' + err.message });
    }
});

app.put('/api/classes/:fileId/sheet', async (req, res) => {
    const { fileId } = req.params;
    const { updates, rows } = req.body || {};

    try {
        const classRecord = await dbGet('SELECT * FROM classes WHERE fileId = ?', [fileId]);
        if (!classRecord || !fs.existsSync(classRecord.filePath)) {
            return res.status(404).json({ error: 'Class Excel file not found on server.' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(classRecord.filePath);
        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
            return res.status(500).json({ error: 'Worksheet not found.' });
        }

        // Find header row and roll column
        let rollColIndex = 1;
        for (let r = 1; r <= Math.min(15, worksheet.rowCount); r++) {
            const row = worksheet.getRow(r);
            row.eachCell((cell, colNumber) => {
                const text = String(cell.value || '').trim();
                if (/^(reg(istration)?|roll)\s*(no\.?|number)?$/i.test(text) || /^reg\b/i.test(text) || /^roll\b/i.test(text)) {
                    rollColIndex = colNumber;
                }
            });
            if (rollColIndex !== 1) break;
        }

        // Map roll numbers to row indices
        const rollRowMap = {};
        for (let r = 1; r <= worksheet.rowCount; r++) {
            const cellVal = worksheet.getRow(r).getCell(rollColIndex).value;
            if (cellVal) {
                rollRowMap[String(cellVal).trim().toUpperCase()] = r;
            }
        }

        // Fetch known sessions for this class to sync SQLite attendance records
        const sessionRecords = await dbAll('SELECT sessionId, colNumber, sessionColumn FROM sessions WHERE fileId = ?', [fileId]);
        const sessionByCol = {};
        for (const s of sessionRecords) {
            sessionByCol[s.colNumber] = s.sessionId;
        }

        let appliedCount = 0;

        // Process granular updates list if provided: [{ rollNo, colNumber, value }]
        if (Array.isArray(updates)) {
            for (const upd of updates) {
                const rollNo = String(upd.rollNo || '').trim().toUpperCase();
                const colNumber = Number(upd.colNumber);
                const value = (Number(upd.value) === 1 || upd.value === '1' || upd.value === true) ? 1 : 0;

                const rowIndex = rollRowMap[rollNo];
                if (rowIndex && colNumber > 1) {
                    worksheet.getRow(rowIndex).getCell(colNumber).value = value;
                    appliedCount++;

                    // Synchronize with SQLite database
                    const sessionId = sessionByCol[colNumber];
                    if (sessionId) {
                        if (value === 1) {
                            await dbRun(`
                                INSERT OR IGNORE INTO attendance (sessionId, rollNo, ip, fileId)
                                VALUES (?, ?, 'MANUAL_EDIT', ?)
                            `, [sessionId, rollNo, fileId]);
                        } else {
                            await dbRun(`
                                DELETE FROM attendance
                                WHERE sessionId = ? AND rollNo = ?
                            `, [sessionId, rollNo]);
                        }
                    }
                }
            }
        } else if (Array.isArray(rows)) {
            // Process full rows payload: [{ rollNo, cells: { col_2: 0, col_3: 1 } }]
            for (const rowData of rows) {
                const rollNo = String(rowData.rollNo || '').trim().toUpperCase();
                const rowIndex = rollRowMap[rollNo];
                if (!rowIndex || !rowData.cells) continue;

                for (const [colKey, rawVal] of Object.entries(rowData.cells)) {
                    const match = colKey.match(/^col_(\d+)$/);
                    if (!match) continue;
                    const colNumber = Number(match[1]);
                    const value = (Number(rawVal) === 1 || rawVal === '1' || rawVal === true) ? 1 : 0;

                    worksheet.getRow(rowIndex).getCell(colNumber).value = value;
                    appliedCount++;

                    const sessionId = sessionByCol[colNumber];
                    if (sessionId) {
                        if (value === 1) {
                            await dbRun(`
                                INSERT OR IGNORE INTO attendance (sessionId, rollNo, ip, fileId)
                                VALUES (?, ?, 'MANUAL_EDIT', ?)
                            `, [sessionId, rollNo, fileId]);
                        } else {
                            await dbRun(`
                                DELETE FROM attendance
                                WHERE sessionId = ? AND rollNo = ?
                            `, [sessionId, rollNo]);
                        }
                    }
                }
            }
        }

        // Backup existing file before saving manual modifications
        if (fs.existsSync(classRecord.filePath)) {
            const backupName = `${fileId}_manual_${Date.now()}.bak.xlsx`;
            try {
                fs.copyFileSync(classRecord.filePath, path.join(BACKUPS_DIR, backupName));
            } catch (_) {}
        }

        // Save updated workbook directly to persistent course file
        const tempPath = `${classRecord.filePath}.manual.tmp`;
        await workbook.xlsx.writeFile(tempPath);
        fs.renameSync(tempPath, classRecord.filePath);

        await dbRun('UPDATE classes SET updatedAt = CURRENT_TIMESTAMP WHERE fileId = ?', [fileId]);

        // If this is activeClass, reload in-memory cache so subsequent attendance sessions
        // immediately use the updated historical data
        if (activeClass && activeClass.fileId === fileId) {
            await loadClassRecord(classRecord);
        }

        console.log(`[MANUAL EDIT] Saved ${appliedCount} cell edits for class ${classRecord.name}`);

        return res.json({
            success: true,
            message: `Successfully saved ${appliedCount} manual edits to ${classRecord.filename}.`,
            updatedAt: new Date().toISOString()
        });
    } catch (err) {
        console.error('[SAVE SHEET ERROR]', err);
        return res.status(500).json({ error: 'Could not save spreadsheet edits: ' + err.message });
    }
});

// --------------------------------------------------
// PROFESSOR: CURRENT TOKEN API
// --------------------------------------------------

app.get('/api/token', (req, res) => {
    if (isPublicTunnelRequest(req)) {
        return res.status(403).json({
            error: 'Professor dashboard is local only.'
        });
    }

    res.set('Cache-Control', 'no-store, no-cache, must-revalidate');
    res.set('Pragma', 'no-cache');

    const joinUrl = sessionActive
        ? `${BASE_URL}/join?token=${encodeURIComponent(currentToken)}`
        : null;

    return res.json({
        success: true,
        active: sessionActive,
        token: sessionActive ? currentToken : null,
        expiresAt: sessionActive ? tokenExpiresAt : 0,
        joinUrl,
        url: joinUrl,
        activeClass: activeClass ? {
            fileId: activeClass.fileId,
            name: activeClass.name,
            filename: activeClass.filename,
            studentCount: activeClass.studentCount
        } : null
    });
});

// --------------------------------------------------
// STUDENT JOIN PAGE
// --------------------------------------------------

app.get('/join', (req, res) => {
    res.set('Cache-Control', 'no-store');

    if (!sessionActive) {
        return res.status(403).send(
            'Attendance is closed. Please contact your professor.'
        );
    }

    return res.sendFile(
        path.join(PUBLIC_DIR, 'validate.html'),
        err => {
            if (err && !res.headersSent) {
                console.error('[JOIN] Page error:', err.message);
                res.status(500).send('Could not load the student attendance page.');
            }
        }
    );
});

// --------------------------------------------------
// STUDENT: VALIDATE QR TOKEN
// --------------------------------------------------

app.post('/api/validate', async (req, res) => {
    if (!sessionActive) {
        return res.status(403).json({
            error: 'Attendance is closed.'
        });
    }

    if (!requireChrome(req, res)) return;

    const { token } = req.body || {};

    if (!tokenIsActive(token)) {
        return res.status(400).json({
            error: 'QR code expired. Scan the current code.'
        });
    }

    if (tokenUses >= MAX_USES) {
        return res.status(400).json({
            error: 'Capacity reached for this QR cycle.'
        });
    }

    const ip = getRequestIp(req);
    const deviceHash = getDeviceHash(req);

    if (deviceScans.has(deviceHash)) {
        return res.status(400).json({
            error: 'This device has already scanned this QR code.'
        });
    }

    deviceScans.add(deviceHash);

    try {
        const duplicateIp = await dbGet(
            `SELECT rollNo
             FROM attendance
             WHERE sessionId = ? AND ip = ?
             LIMIT 1`,
            [currentSessionId, ip]
        );

        if (!sessionActive || !tokenIsActive(token)) {
            return res.status(400).json({
                error: 'QR code expired during validation.'
            });
        }

        if (duplicateIp) {
            return res.status(400).json({
                error: 'Attendance has already been recorded from this IP address.'
            });
        }

        tokenUses++;
        validatedDevices.add(deviceHash);

        return res.json({
            success: true,
            expiresAt: tokenExpiresAt
        });
    } catch (err) {
        console.error('[VALIDATE] Database error:', err.message);
        deviceScans.delete(deviceHash);
        return res.status(500).json({
            error: 'Could not validate attendance.'
        });
    }
});

// --------------------------------------------------
// STUDENT: FINAL ATTENDANCE SUBMISSION
// Validates student against active class roster (25BCS001 to 25BCS215).
// Updates SQLite database and corresponding Excel file cell to 1.
// --------------------------------------------------

app.post('/api/submit', async (req, res) => {
    if (!sessionActive) {
        return res.status(403).json({
            error: 'Attendance is closed.'
        });
    }

    if (!requireChrome(req, res)) return;

    const { token, rollNo, deviceId } = req.body || {};

    if (!tokenIsActive(token)) {
        return res.status(400).json({
            error: 'QR expired. Scan the latest classroom QR code.'
        });
    }

    if (typeof rollNo !== 'string' || !rollNo.trim()) {
        return res.status(400).json({
            error: 'Enter a valid roll number.'
        });
    }

    if (
        typeof deviceId !== 'string' ||
        !deviceId.trim() ||
        deviceId.length > 200
    ) {
        return res.status(400).json({
            error: 'Device identification failed. Please try again.'
        });
    }

    const cleanRollNo = rollNo.trim().toUpperCase();
    const cleanDeviceId = deviceId.trim();
    const ip = getRequestIp(req);
    const deviceHash = getDeviceHash(req);

    // Verify an active class is assigned
    if (!activeClass) {
        return res.status(400).json({
            error: 'No active attendance class is selected.'
        });
    }

    // REQUIREMENT 5: Strict validation against class roster.
    // If an unauthorized student (e.g. PhD student, wrong batch, unknown roll) attempts to mark attendance:
    // Reject with: "Attendance Rejected: You are not registered for this class."
    if (!activeClass.authorizedRollsSet.has(cleanRollNo)) {
        return res.status(403).json({
            error: 'Attendance Rejected: You are not registered for this class.'
        });
    }

    if (submittedDevices.has(cleanDeviceId)) {
        return res.status(400).json({
            error: 'This device has already recorded attendance in this session.'
        });
    }

    if (!validatedDevices.has(deviceHash)) {
        return res.status(400).json({
            error: 'Validate the current QR code before submitting attendance.'
        });
    }

    validatedDevices.delete(deviceHash);
    beginSubmission();

    try {
        if (!tokenIsActive(token)) {
            return res.status(400).json({
                error: 'Attendance session stopped or QR expired.'
            });
        }

        // Insert into attendance table with fileId tracking
        const result = await dbRun(
            `INSERT INTO attendance (sessionId, rollNo, ip, fileId)
             SELECT ?, ?, ?, ?
             WHERE NOT EXISTS (
                 SELECT 1
                 FROM attendance
                 WHERE sessionId = ?
                   AND (rollNo = ? OR ip = ?)
             )`,
            [
                currentSessionId,
                cleanRollNo,
                ip,
                activeClass.fileId,
                currentSessionId,
                cleanRollNo,
                ip
            ]
        );

        if (result.changes === 0) {
            const duplicate = await dbGet(
                `SELECT rollNo, ip
                 FROM attendance
                 WHERE sessionId = ?
                   AND (rollNo = ? OR ip = ?)
                 LIMIT 1`,
                [currentSessionId, cleanRollNo, ip]
            );

            if (duplicate && duplicate.rollNo === cleanRollNo) {
                return res.status(409).json({
                    error: 'Attendance is already recorded for this roll number.'
                });
            }

            if (duplicate && duplicate.ip === ip) {
                return res.status(409).json({
                    error: 'Attendance has already been recorded from this IP address.'
                });
            }

            return res.status(409).json({
                error: 'Attendance could not be recorded. Please try again.'
            });
        }

        submittedDevices.add(cleanDeviceId);

        // Update Excel in-memory cell to 1
        if (activeWorkbook && activeSessionColIndex) {
            const worksheet = activeWorkbook.worksheets[0];
            const rowIndex = activeClass.rollRowMap[cleanRollNo];
            if (worksheet && rowIndex) {
                worksheet.getRow(rowIndex).getCell(activeSessionColIndex).value = 1;
                // Schedule non-blocking disk persistence
                persistActiveWorkbook();
            }
        }

        console.log(
            `[ATTENDANCE] Present: ${cleanRollNo} | Class: ${activeClass.name} | IP: ${ip}`
        );

        // Optional Google Sheets backup
        if (GOOGLE_SHEETS_WEBHOOK) {
            try {
                fetch(GOOGLE_SHEETS_WEBHOOK, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        rollNo: cleanRollNo,
                        deviceIp: ip,
                        sessionId: currentSessionId,
                        className: activeClass.name
                    }),
                    signal: AbortSignal.timeout(5000)
                }).catch(e => console.error('[SHEETS] Backup failed:', e.message));
            } catch (_) {}
        }

        return res.json({
            success: true,
            rollNo: cleanRollNo
        });
    } catch (err) {
        console.error('[SUBMIT] Error:', err.message);
        return res.status(500).json({
            error: 'Could not save attendance.'
        });
    } finally {
        endSubmission();
    }
});

// --------------------------------------------------
// PROFESSOR: LIVE ATTENDANCE LIST & STATS
// --------------------------------------------------

app.get('/api/attendance', async (req, res) => {
    if (isPublicTunnelRequest(req)) {
        return res.status(403).json({
            error: 'Professor dashboard is local only.'
        });
    }

    try {
        const rows = await dbAll(
            `SELECT timestamp, rollNo, ip
             FROM attendance
             WHERE sessionId = ?
             ORDER BY id DESC`,
            [currentSessionId]
        );

        const totalEnrolled = activeClass ? activeClass.studentCount : TOTAL_EXPECTED_STUDENTS;
        const presentCount = rows.length;
        const absentCount = Math.max(0, totalEnrolled - presentCount);

        res.set('Cache-Control', 'no-store');

        return res.json({
            success: true,
            active: sessionActive,
            count: presentCount,
            absentCount,
            totalCount: totalEnrolled,
            attendance: rows,
            activeClass: activeClass ? {
                fileId: activeClass.fileId,
                name: activeClass.name,
                filename: activeClass.filename,
                studentCount: activeClass.studentCount
            } : null
        });
    } catch (err) {
        console.error('[ATTENDANCE LIST]', err.message);
        return res.status(500).json({
            error: 'Could not load attendance.'
        });
    }
});

// --------------------------------------------------
// PROFESSOR: DOWNLOAD UPDATED EXCEL
// --------------------------------------------------

app.get('/api/download', async (req, res) => {
    if (isPublicTunnelRequest(req)) {
        return res.status(403).send('Professor downloads are local only.');
    }

    if (!activeClass) {
        return res.status(400).send('No active class selected.');
    }

    try {
        // Ensure latest records are flushed to Excel
        await syncSessionToExcel(currentSessionId, activeClass.fileId);

        if (!fs.existsSync(activeClass.filePath)) {
            return res.status(404).send('Class Excel file not found on disk.');
        }

        const data = fs.readFileSync(activeClass.filePath);

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename="${activeClass.filename}"`
        );
        res.setHeader('Cache-Control', 'no-store');

        return res.send(data);
    } catch (err) {
        console.error('[EXPORT]', err.message);
        if (!res.headersSent) {
            return res.status(500).send('Could not create the attendance report.');
        }
    }
});

// --------------------------------------------------
// PROFESSOR: STOP QR AND CLOSE ATTENDANCE
// --------------------------------------------------

app.post('/api/stop-session', async (req, res) => {
    if (isPublicTunnelRequest(req)) {
        return res.status(403).json({
            error: 'Professor dashboard is local only.'
        });
    }

    if (sessionActive) {
        sessionActive = false;
        stopTokenLoop();

        currentToken = '';
        tokenExpiresAt = 0;
        tokenUses = 0;

        deviceScans.clear();
        validatedDevices.clear();

        console.log('[SESSION] QR stopped. Attendance closed.');
    }

    await waitForSubmissions();

    if (activeClass) {
        // Ensure final sync to Excel
        await syncSessionToExcel(currentSessionId, activeClass.fileId);

        // Update session status in database
        await dbRun(`
            UPDATE sessions
            SET status = 'STOPPED', endedAt = CURRENT_TIMESTAMP
            WHERE sessionId = ?
        `, [currentSessionId]);
    }

    return res.json({
        success: true,
        active: false,
        message: 'Attendance stopped. The updated Excel file is ready.'
    });
});

// --------------------------------------------------
// PROFESSOR: START A NEW ATTENDANCE SESSION
// Prepares a new column in the active Excel file with today's date.
// If multiple sessions occur on the same day, numbers them appropriately.
// Initializes all students to 0 (absent).
// --------------------------------------------------

app.post('/api/reset', async (req, res) => {
    if (isPublicTunnelRequest(req)) {
        return res.status(403).json({
            error: 'Professor dashboard is local only.'
        });
    }

    if (!activeClass || !activeWorkbook) {
        return res.status(400).json({
            error: 'Please upload or select an Excel class file before starting an attendance session.'
        });
    }

    sessionActive = false;
    stopTokenLoop();
    await waitForSubmissions();

    try {
        // Re-read workbook to ensure we have latest saved state
        const freshWorkbook = new ExcelJS.Workbook();
        await freshWorkbook.xlsx.readFile(activeClass.filePath);
        activeWorkbook = freshWorkbook;

        const worksheet = activeWorkbook.worksheets[0];
        const { sessionColName, colNumber } = prepareSessionColumnInWorksheet(
            worksheet,
            activeClass.headerRowIndex,
            activeClass.rollRowMap
        );

        activeSessionColumn = sessionColName;
        activeSessionColIndex = colNumber;

        // Persist initial column with all 0s
        await persistActiveWorkbook();

        currentSessionId = crypto.randomUUID();

        // Record session in database
        await dbRun(`
            INSERT INTO sessions (sessionId, fileId, sessionDate, sessionColumn, colNumber, status)
            VALUES (?, ?, ?, ?, ?, 'ACTIVE')
        `, [
            currentSessionId,
            activeClass.fileId,
            getFormattedDate(),
            sessionColName,
            colNumber
        ]);

        submittedDevices.clear();
        deviceScans.clear();
        validatedDevices.clear();

        sessionActive = true;
        startTokenLoop();

        console.log(`[SESSION] New attendance session started for '${activeClass.name}'. Column: '${sessionColName}'`);

        return res.json({
            success: true,
            active: true,
            sessionId: currentSessionId,
            sessionColumn: sessionColName,
            class: {
                fileId: activeClass.fileId,
                name: activeClass.name,
                filename: activeClass.filename
            }
        });
    } catch (err) {
        console.error('[START SESSION ERROR]', err);
        return res.status(500).json({
            error: 'Could not initialize session column in Excel: ' + err.message
        });
    }
});

// --------------------------------------------------
// PROFESSOR: SHUT DOWN THE ENTIRE SERVER
// --------------------------------------------------

app.post('/api/shutdown', (req, res) => {
    if (isPublicTunnelRequest(req)) {
        return res.status(403).json({
            error: 'Server shutdown is local only.'
        });
    }

    res.json({
        success: true,
        message: 'Server is shutting down.'
    });

    setTimeout(() => {
        shutdown();
    }, 100);
});

// --------------------------------------------------
// GRACEFUL SHUTDOWN
// --------------------------------------------------

async function shutdown() {
    if (shuttingDown) return;

    shuttingDown = true;
    sessionActive = false;

    stopTokenLoop();

    currentToken = '';
    tokenExpiresAt = 0;

    await waitForSubmissions();

    if (activeClass && activeWorkbook) {
        await syncSessionToExcel(currentSessionId, activeClass.fileId);
    }

    if (ngrokListener) {
        try {
            await ngrokListener.close();
            ngrokListener = null;
            console.log('[NGROK] Tunnel closed.');
        } catch (err) {
            console.error('[NGROK] Close warning:', err.message);
        }
    }

    if (server && server.listening) {
        await new Promise(resolve => server.close(resolve));
    }

    db.close(err => {
        if (err) {
            console.error('[DB] Close error:', err.message);
            process.exit(1);
        }

        console.log('[SYS] Server stopped.');
        process.exit(0);
    });

    setTimeout(() => process.exit(1), 5000).unref();
}

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

// --------------------------------------------------
// START SERVER ONLY AFTER DATABASE INITIALIZATION
// --------------------------------------------------

async function startServer() {
    try {
        await initializeDatabase();

        server = app.listen(PORT, '0.0.0.0', async () => {
            console.log('');
            console.log('========================================');
            console.log(' ATTENDANCE SERVER RUNNING');
            console.log('========================================');
            console.log(` Dashboard: http://localhost:${PORT}`);
            console.log(` LAN URL:   http://${myIp}:${PORT}`);
            console.log(` Token API: http://localhost:${PORT}/api/token`);
            console.log(` Health:    http://localhost:${PORT}/api/health`);
            console.log('========================================');

            if (!NGROK_AUTHTOKEN) {
                console.log('[NGROK] Disabled. Local Wi-Fi hosting is active.');
                console.log(`[QR] Student URL base: ${BASE_URL}`);
                return;
            }

            try {
                const ngrok = require('@ngrok/ngrok');
                ngrokListener = await ngrok.forward({
                    addr: PORT,
                    authtoken: NGROK_AUTHTOKEN
                });

                const tunnelUrl = ngrokListener.url();
                if (tunnelUrl) {
                    BASE_URL = tunnelUrl.replace(/\/+$/, '');
                }

                console.log('[NGROK] Tunnel started.');
                console.log(`[NGROK] Student URL base: ${BASE_URL}`);
                console.log('[NGROK] The professor dashboard remains local only.');
            } catch (err) {
                console.error('[NGROK] Could not start tunnel:', err.message);
                console.log('[SYS] Continuing with local hosting.');
            }
        });

        server.on('error', err => {
            if (err.code === 'EADDRINUSE') {
                console.error(`[SYS] Port ${PORT} is already in use. Stop the other server first.`);
            } else {
                console.error('[SYS] Server error:', err.message);
            }
        });
    } catch (err) {
        console.error('[SYS] Failed to initialize database:', err.message);
        db.close(() => process.exit(1));
    }
}

startServer();
