/**
 * Fetch Bank Payment + Receipt Vouchers from Tally for all loaded companies (Apr–Jun 2026)
 * Outputs one consolidated Excel workbook:
 *   - Summary              : per-company payment & receipt totals
 *   - ALL                  : pivot by Final Voucher / BS/PL / TYPE — all transactions
 *   - BANK / CASH / RETUNN : same pivot filtered by Voucher Classification
 *   - Related Partry       : Category=Related party breakdown
 *   - Bracnch Mis          : Category=Branch party breakdown
 *   - All Transactions     : flat ledger, all companies combined
 *   - <Company>            : one sheet per company
 *
 * Classification (payment):  positive (credit) = Bank Account; all negatives = Party rows
 * Classification (receipt):  negative + bank-group ancestor = Bank Account; everything else = Party rows
 *
 * Usage:  node fetch-bank-transactions.mjs
 */

import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { XMLParser } from 'fast-xml-parser';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

const __dirname  = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_DIR = path.join(__dirname, '..', '..', '..', '..');
const OUTPUT_DIR = path.join(PROJECT_DIR, 'outputs');

// Load .env the same way the Python scripts do: CWD first (project root when run
// as documented), then the project dir resolved from this file's location, so the
// script works from any working directory. Real environment variables always win.
for (const envPath of [path.join(process.cwd(), '.env'), path.join(PROJECT_DIR, '.env')]) {
    if (!fs.existsSync(envPath)) continue;
    for (const line of fs.readFileSync(envPath, 'utf8').split(/\r?\n/)) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) continue;
        const eq = trimmed.indexOf('=');
        if (eq === -1) continue;
        const key = trimmed.slice(0, eq).trim();
        if (process.env[key] === undefined) process.env[key] = trimmed.slice(eq + 1).trim();
    }
    break;
}

const TALLY_HOST      = process.env.TALLY_HOST || 'localhost';
const TALLY_PORT      = parseInt(process.env.TALLY_PORT  || '9000');
const FROM_DATE       = process.env.FROM_DATE       || '1-Apr-2026';
const TO_DATE         = process.env.TO_DATE         || '4-Aug-2026';
const OUTPUT_FILE     = process.env.OUTPUT_FILE     || path.join(OUTPUT_DIR, 'Bank_Transactions_Apr_Aug2026.xlsx');
const REFERENCE_FILE  = process.env.REFERENCE_FILE  || path.join(OUTPUT_DIR, 'Bank_Transactions_Apr_Jun_2026_v3.xlsx');
const EXPENSE_FILE    = process.env.EXPENSE_FILE    || path.join(OUTPUT_DIR, '..', 'Files', 'EXPENSES REFERESH', 'Expense Sheet 2026-07-14 17-38-44.xlsx');
const COMPANY_LIST_FILE = process.env.COMPANY_LIST_FILE || path.join(OUTPUT_DIR, '..', 'Files', 'TALLY Company_List.xlsx');

// --period 2627 | 2526 | 2425  (also settable via COMPANY_CODE env var). Defaults to 26-27.
// See Files/TALLY Company_List.xlsx and CLAUDE.md "Company Period Filtering".
function getPeriodArg() {
    const idx = process.argv.indexOf('--period');
    if (idx !== -1 && process.argv[idx + 1]) return process.argv[idx + 1];
    return process.env.COMPANY_CODE || '26-27';
}
const COMPANY_CODE = getPeriodArg();

const MAX_LED = 25;

// Company → short code + location (matches user's manual classification in the Pivot file).
// Keyed by every known Tally company-FILE name across all 3 tracked periods
// (26-27/25-26/24-25, per Files/TALLY Company_List.xlsx) so labeling stays
// correct no matter which --period was fetched, not just the current one.
const COMPANY_META = new Map([
    // Otec-Surat
    ['ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)',         { coCode: 'OTEC ',    location: 'ST' }],
    ['ORANGE O TEC PRIVATE LIMITED (01-04-24TO31-03-25)',         { coCode: 'OTEC ',    location: 'ST' }],
    // Otec-Noida
    ['ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)',        { coCode: 'OTEC ',    location: 'ST' }],
    ['ORANGE O TEC PRIVATE LIMITED-NOIDA - (from 1-Apr-24)',      { coCode: 'OTEC ',    location: 'ST' }],
    // Ent-Surat
    ['ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)',             { coCode: 'ENTER ',   location: 'ST' }],
    ['ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26)',             { coCode: 'ENTER ',   location: 'ST' }],
    // Ent-Noida
    ['ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27',  { coCode: 'ENTER',    location: 'NOIDA' }],
    ['ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26',  { coCode: 'ENTER',    location: 'NOIDA' }],
    // Colorix
    ['COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)',  { coCode: 'COLORIX',  location: 'ST' }],
]);

// ---------------------------------------------------------------------------
// Bank-group hierarchy check (for receipt classification)
// ---------------------------------------------------------------------------

function isLedgerInBankGroup(ledgerParent, groupMap) {
    if (!ledgerParent) return false;
    let current = ledgerParent;
    const visited = new Set();
    while (current) {
        if (visited.has(current)) return false;
        visited.add(current);
        const l = current.toLowerCase().trim();
        if (l === 'bank accounts' || l === 'bank od account' || l === 'bank od a/c') return true;
        if (l === 'primary' || l === '') return false;
        current = groupMap.get(current) || '';
    }
    return false;
}

// ---------------------------------------------------------------------------
// Row classification — 6 computed columns
// ---------------------------------------------------------------------------

const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const MONTH_ORDER  = MONTHS_SHORT; // same array, used for sorting

function dateToMonth(date) {
    if (!date) return '';
    const parts = date.split('-');
    if (parts.length !== 3) return '';
    const m = parseInt(parts[1], 10);
    const y = parts[2].slice(2);
    if (m < 1 || m > 12) return '';
    return MONTHS_SHORT[m - 1] + '-' + y;
}

function isRelatedParty(party) {
    const p = (party || '').toUpperCase();
    return p.includes('ORANGE O TEC') || p.includes('KARAN IMPEX') || p.includes('KARAN TOSHNIWAL');
}

// Direct Sub Group → TYPE mapping (covers ~95% of rows)
const SUBGROUP_TYPE = new Map([
    // Current Liabilities — creditors & payables
    ['ADVANCE FOR CAPITAL GOODS',            'FA'],
    ['BALANCE WITH RELATED PARTY',           'Related'],
    ['BANK CHARGES',                         'FINANCE COST'],
    ['BANK CHARGES-IMPORT DIRECT PAYMENT',   'FINANCE COST'],
    ['BANK CHARGES-LC-BC',                   'FINANCE COST'],
    ['BANK LC-BC LIMIT',                     'BC Loan'],
    ['CARRIAGE INWARD EXP',                  'Carriage Inward'],
    ['CARRIAGE OUTWARD ACCOUNTS',            'Carriage Outward'],
    ['CARRIAGE OUTWARD EXP',                 'Carriage Outward'],
    ['CORPORATION COMPANIES-TDS',            'TDS'],
    ['COURIER & POSTAGE EXP',                'POSTAGE'],
    ['CREDIT CARD',                          'Credit Card'],
    ['CREDITOR FOR EMPLOYEE (EXP)',          'EMPLOYEE TRAVEL COST'],
    // CREDITOR FOR EMPLOYEE handled in getType() — Salary vs Travel by party name
    ['CREDITOR FOR GOODS',                   'Local Goods'],
    ['CREDITOR FOR GOODS-IMPORT',            'IMPORT GOOD'],
    ['CREDITOR FOR GOODS-LOCAL_MSME',        'Local Goods'],
    ['CREDITOR FOR GOODS-LOCAL_NON MSME',    'Local Goods'],
    ['CREDITOR FOR GOODS-OTHER_MSME',        'Local Goods'],
    ['CREDITOR FOR GOODS-OTHER_NON MSME',    'Local Goods'],
    // CREDITOR FOR OTHER* handled in getType() — party-name patterns

    ['CREDITORS FOR COMMISSION',             'COMMISSION & BROKERAGE EXP'],
    ['CREDITORS FOR COMMISSION-NON MSME',    'COMMISSION & BROKERAGE EXP'],
    ['CREDITORS FOR GOODS',                  'Local Goods'],
    ['CREDITORS FOR GOODS-MSME',             'Local Goods'],
    ['CREDITORS FOR GOODS-NON MSME',         'Local Goods'],
    ['CREDITORS FOR OTHER',                  'EXPENSES'],
    ['CREDITORS FOR OTHER-MSME',             'EXPENSES'],
    // CREDITORS FOR OTHER-NON MSME handled in getType() — party-name patterns

    ['CUSTOM DUTY ACCOUNT',                  'Duty'],
    ['CUSTOM DUTY EXPENSES',                 'Duty'],
    ['EMPLOYEE SALARY EXPENSES',             'EMPLOYEE COST'],
    ['ESIC, PF & PT PAYABLE',               'PF ESIC'],
    ['FINANCE COST',                         'FINANCE COST'],
    ['GST ACCOUNT',                          'GST'],
    ['HOTEL EXPENSES',                       'EXPENSES'],
    ['INTEREST COST',                        'FINANCE COST'],
    ['LEGAL AND PROFESSIONAL EXP',           'EXPENSES'],
    ['LOADING & UNLOADING EXP',              'EXPENSES'],
    ['OFFICE EXPENSES',                      'EXPENSES'],
    // OTHER BANK CHARGES handled in getType() — Bank Charges vs EXPENSES by party name

    ['OTHER THEN COMPANIES-TDS',             'TDS'],
    ['PRINTING STATIONERY EXP',              'EXPENSES'],
    ['REPAIR & MAINTENANCE COST',            'EXPENSES'],
    ['REPAIR & MAINTENANCE EXP (OFFICE)',    'EXPENSES'],
    ['SERVICE EXP',                          'EXPENSES'],
    ['STAFF WELFARE EXPENSES',               'EXPENSES'],
    ['SUBSCRIPTION & MEMBERSHIP EXP',        'EXPENSES'],
    ['SUNDRY CREDITORS',                     'EXPENSES'],
    ['TAXES EXP',                            'Stamp Paper'],
    ['TDS-TCS PAYABLE',                      'TDS'],
    ['TDS-TCS Payable-ISD',                  'TDS'],
    ['TOOLS & SPARE EXPENSE',                'EXPENSES'],
    ['TRAVELLING EXP',                       'EXPENSES'],
    ['VEHICAL EXP',                          'EXPENSES'],
    // Current Assets — debtors
    ['Advance Tax',                          'Tax'],
    ['Ahmedabad',                            'Debtors'],
    ['Amritsar',                             'Debtors'],
    ['Bangalore',                            'Debtors'],
    ['EMPLOYEE ADVANCES ACCOUNTS',           'EMP ADV'],
    ['Faridabad',                            'Debtors'],
    ['Gujarat',                              'Debtors'],
    ['Haryana',                              'Debtors'],
    ['Hosur',                                'Debtors'],
    ['Jodhpur',                              'Debtors'],
    ['Karnataka',                            'Debtors'],
    ['Karur',                                'Debtors'],
    ['Kolhapur',                             'Debtors'],
    ['Ludhiana',                             'Debtors'],
    ['MACHINE DEBTORS',                      'Debtors'],
    ['Maharashtra',                          'Debtors'],
    ['Mumbai',                               'Debtors'],
    ['NAKUL SALE',                           'Debtors'],
    ['OTHER DEBTORS',                        'Debtors'],
    ['Punjab',                               'Debtors'],
    ['Rajasthan',                            'Debtors'],
    ['SalesMan',                             'Debtors'],
    ['Sundry Debtors',                       'Debtors'],
    ['Surat',                                'Debtors'],
    ['Surat (Nakul)',                        'Debtors'],
    ['Surat(Umesh)',                         'Debtors'],
    ['TDS PAID AMT RECEIVABLE',              'TDS'],
    ['TELANGANA',                            'Debtors'],
    ['Tirupur',                              'Debtors'],
    ['Varanasi',                             'Debtors'],
    // Investments
    ['DYNAMIC ASSET ALLOCATION OR BALANCED ADVANTAGE', 'INV'],
    ['FLEXI CAP FUND',                       'INV'],
    ['Fixed Deposit',                        'INV'],
    ['MULTI ASSET ALLOCATION',               'INV'],
    // Fixed Assets
    ['Fixed Assets',                         'FA'],
    // Suspense
    ['Suspense A/c',                         'Suspense A/c'],
    // Branch
    ['Branch / Divisions',                   'Branch'],
]);

// TYPE → BS/PL
const TYPE_BSPL = {
    // BS
    'Debtors':        'BS', 'IMPORT GOOD':  'BS', 'Local Goods':  'BS', 'FA':          'BS',
    'INV':            'BS', 'EMP ADV':      'BS', 'TDS':          'BS', 'GST':         'BS',
    'PF ESIC':        'BS', 'Related':      'BS', 'Branch':       'BS', 'Suspense A/c':'BS',
    'Tax':            'BS',
    // PL — legacy short names (kept for backwards compatibility)
    'EXPENSES':       'PL', 'Travel':       'PL', 'Commission':   'PL', 'Credit Card':  'PL',
    'Duty':           'PL', 'Bank Charges': 'PL', 'Salary':       'PL', 'Transport':   'PL',
    'Freight Charges':'PL', 'Marketing':    'PL', 'Rent':         'PL', 'ELECTRICITY': 'PL',
    'Seurity':        'PL', 'Carriage Inward':'PL','Carriage Outward':'PL','POSTAGE':  'PL',
    'Postage':        'PL', 'Stamp Duty':   'PL',
    // PL — granular expense TYPEs (from user v3 classification)
    'FINANCE COST':              'PL', 'EMPLOYEE COST':             'PL',
    'EMPLOYEE TRAVEL COST':      'PL', 'TRANSPORTATION EXP':        'PL',
    'CLEARING & FORWERDING EXP': 'PL', 'ELECTRICITY EXPENSES':      'PL',
    'RENT EXPENSES':             'PL', 'REPAIR & MAINTENANCE COST': 'PL',
    'ADVERTISING PROMOTION EXP': 'PL', 'LEGAL AND PROFESSIONAL EXP':'PL',
    'COMMISSION & BROKERAGE EXP':'PL', 'COMPUTER & PRINTER EXP':    'PL',
    'PRINTING AND COURIER':      'PL', 'PACKING EXP':               'PL',
    'SUBSCRIPTION & MEMBERSHIP EXP':'PL','INSURANCE EXP':           'PL',
    'LABORATORY EXPENCES':       'PL', 'LABOUR COST':               'PL',
    'SERVICE EXP':               'PL', 'SAMPLE FABRIC EXPENSES':    'PL',
    'Software':                  'PL', 'SECURITY EXP':              'PL',
    'ENVIRONMENTAL PROTECTION':  'PL', 'DONATION SUBSCRIPTION':     'PL',
    'Learning and Development':  'PL', 'Other Diretc Expenses':     'PL',
    'Stamp Paper':               'PL', 'CARRIAGE OUTWARD EXP':      'PL',
    'CARRIAGE INWARD EXP':       'PL',
    // Loan
    'BC Loan':        'Loan', 'Car Loan':   'Loan', 'ECLGS Loan': 'Loan',
    'TERM LOAN':      'Loan', 'Term Loan':  'Loan', 'OTH LOAN':   'Loan',
};

// Hardcoded overrides for data-quality issues (wrong ledger group in Tally)
const PARTY_OVERRIDES = new Map([
    ["MITHILI PROFFESIONAL'S-Cross Creation", { type: 'Debtors', category: 'OTH', bspl: 'BS' }],
]);

// Vouchers entered as plain Payment in Tally whose narration confirms they are actually
// bounced-cheque returns (bank REJECT/RTN codes) — force Voucher Classification to CHQ.Ret.
// Found via narration scan of "All Transactions" for RTN/REJECT/RETURN not already tagged CHQ.Ret.
const CHQ_RETURN_GUID_OVERRIDES = new Set([
    '393ee4bd-4fdc-4aed-ae88-e2ff1394927a-00000e8c', // COLORIX / RIDDHIMAN MILLS TEXTILE PRINTING / 35,896
    '779c26f4-3fd8-46bd-9995-4f9916c98856-00001455', // ORANGE O TEC ENTERPRISES / EKTUHI CREATOR / 25-May / 235,000
    '53d35745-5246-4e1a-a27a-d4769f245b50-00004c87', // ORANGE O TEC PRIVATE / WESTEND ENTERPRISES PVT LTD / 2,950
    '59a6c2d9-0c5a-4fc5-b8c5-3be6fec3289e-00003b8d', // ORANGE O TEC ENTERPRISES / PANACHE CREATION / 16-Jun / 100,000
    '779c26f4-3fd8-46bd-9995-4f9916c98856-000015d2', // ORANGE O TEC ENTERPRISES / EKTUHI CREATOR / 19-Jun / 235,000
    '779c26f4-3fd8-46bd-9995-4f9916c98856-000015d4', // ORANGE O TEC ENTERPRISES / EKTUHI CREATOR / 20-Jun / 235,000
]);

function getType(subGroup, party) {
    const sg = (subGroup || '').trim();
    const p  = party || '';

    // Secured Loans: car / ECLGS / term loan by party name
    if (sg === 'Secured Loans') {
        if (/ECLGS/i.test(p))
            return 'ECLGS Loan';
        if (/CAR|BOLERO|KUSHAQ|TAIGUN|VITARA|JIMNY|SELTOS|EV6|WINDOSER/i.test(p))
            return 'Car Loan';
        return 'TERM LOAN';
    }

    if (sg === 'Unsecured Loans')
        return isRelatedParty(p) ? 'Related' : 'OTH LOAN';

    if (sg === 'BALANCE WITH RELATED PARTY(Debtors)')
        return isRelatedParty(p) ? 'Related' : 'Local Goods';

    if (sg === 'Loans & Advances (Asset)' || sg === 'Current Assets')
        return isRelatedParty(p) ? 'Related' : 'Debtors';

    if (sg === 'Sundry Creditors' || sg.toUpperCase() === 'SUNDRY CREDITORS')
        return isRelatedParty(p) ? 'Related' : 'EXPENSES';

    // CREDITOR FOR EMPLOYEE (EXP): default for unknown parties
    if (sg === 'CREDITOR FOR EMPLOYEE (EXP)')
        return 'EMPLOYEE TRAVEL COST';

    // CREDITOR FOR EMPLOYEE (plain): SALARY → EMPLOYEE COST, (EXP) → TRANSPORTATION EXP, else → EMPLOYEE COST
    if (sg === 'CREDITOR FOR EMPLOYEE') {
        if (/SALARY/i.test(p))  return 'EMPLOYEE COST';
        if (/\(EXP\)/i.test(p)) return 'TRANSPORTATION EXP';
        return 'EMPLOYEE COST';
    }

    // CREDIT CARD: most entries are employee travel; default to EMPLOYEE TRAVEL COST
    if (sg === 'CREDIT CARD')
        return 'EMPLOYEE TRAVEL COST';

    // CREDITOR FOR OTHER* and CREDITORS FOR OTHER*: reference file handles specifics;
    // pattern-based fallback for unknown parties
    if (sg === 'CREDITOR FOR OTHER' || sg === 'CREDITOR FOR OTHER- NON MSME' ||
        sg === 'CREDITOR FOR OTHER-MSME' || sg === 'CREDITORS FOR OTHER-NON MSME' ||
        sg === 'CREDITORS FOR OTHER' || sg === 'CREDITORS FOR OTHER-MSME') {
        if (/\bCLEARING\b|FORWARDING|FORWERDING|LOGISTICS|SHIPPING|^DHL\b|NHAVA SEVA|DISTRIPARKS/i.test(p))
            return 'CLEARING & FORWERDING EXP';
        if (/\bTRANSPORT\b|\bROADWAYS\b|AUTO DRIVER/i.test(p))
            return 'TRANSPORTATION EXP';
        if (/CARRIAGE/i.test(p))
            return 'CARRIAGE INWARD EXP';
        if (/\(ELE\)|DGVCL|VIJ COMPANY|ELECTRICITY|ELECTRIC\b/i.test(p))
            return 'ELECTRICITY EXPENSES';
        if (/\bRENT\b|\(RENT\)/i.test(p))
            return 'RENT EXPENSES';
        if (/INSURANCE/i.test(p))
            return 'INSURANCE EXP';
        if (/SECURITY|SEURITY/i.test(p))
            return 'SECURITY EXP';
        if (/HOTEL|RESIDENCY|INN\b|BANQUET|PLAZA/i.test(p))
            return 'EMPLOYEE TRAVEL COST';
        if (/TRAVEL|HOLIDAY|TOUR\b|IRCTC|AVIATION/i.test(p))
            return 'EMPLOYEE TRAVEL COST';
        if (/REPAIR|MAINTENANCE|ELECTRICAL WORKS/i.test(p))
            return 'REPAIR & MAINTENANCE COST';
        return 'EXPENSES';
    }

    // OTHER BANK CHARGES: loan processing charges → FINANCE COST, rest → EXPENSES
    if (sg === 'OTHER BANK CHARGES')
        return /LOAN/i.test(p) ? 'FINANCE COST' : 'EXPENSES';

    return SUBGROUP_TYPE.get(sg) || '';
}

// Read the manually-classified reference file and build a party→{type,category,bspl} map.
// Columns (0-indexed): 6=Party, 15=Category, 16=TYPE, 17=BS/PL
async function loadPartyTypeReference(refFile) {
    const map = new Map();
    if (!fs.existsSync(refFile)) {
        console.log('  No reference file found — using pattern-based classification only');
        return map;
    }
    try {
        const wb   = XLSX.readFile(refFile);
        const ws   = wb.Sheets['All Transactions'];
        if (!ws) return map;
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        for (let r = 1; r < rows.length; r++) {
            const party = String(rows[r][6]  || '').trim();
            const cat   = String(rows[r][15] || '').trim();
            const type  = String(rows[r][16] || '').trim();
            const bspl  = String(rows[r][17] || '').trim();
            if (party && type && !map.has(party))
                map.set(party, { type, category: cat, bspl });
        }
        console.log(`  Reference file: ${map.size} party→TYPE mappings loaded`);
    } catch (err) {
        console.warn(`  Warning: Could not read reference file: ${err.message}`);
    }
    return map;
}

// Read the Expenses workbook's RAW DATA sheet and build a Ledger Name → MAIN GROUP
// map (most-frequent MAIN GROUP per ledger, since some ledgers span multiple groups
// across different expense postings). Returns { map, ambiguous } where ambiguous
// lists ledgers that mapped to more than one distinct MAIN GROUP.
function loadExpenseMainGroupRef(expenseFile) {
    const map = new Map();
    const ambiguous = [];
    if (!fs.existsSync(expenseFile)) {
        console.log('  No Expense Sheet found — "Expense Main Group" column will be blank');
        return { map, ambiguous };
    }
    try {
        const wb = XLSX.readFile(expenseFile);
        const ws = wb.Sheets['RAW DATA'];
        if (!ws) return { map, ambiguous };
        const rows  = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        const header = rows[0].map((h) => String(h || '').trim());
        const lIdx = header.indexOf('LEDGER NAME');
        const mIdx = header.indexOf('MAIN GROUP');
        if (lIdx === -1 || mIdx === -1) return { map, ambiguous };

        const counts = new Map(); // ledger -> Map(mainGroup -> count)
        for (let r = 1; r < rows.length; r++) {
            const ledger = String(rows[r][lIdx] || '').trim();
            const mg     = String(rows[r][mIdx] || '').trim();
            if (!ledger || !mg) continue;
            if (!counts.has(ledger)) counts.set(ledger, new Map());
            const inner = counts.get(ledger);
            inner.set(mg, (inner.get(mg) || 0) + 1);
        }
        for (const [ledger, inner] of counts) {
            const sorted = [...inner.entries()].sort((a, b) => b[1] - a[1]);
            map.set(ledger, sorted[0][0]);
            if (sorted.length > 1) ambiguous.push({ ledger, groups: sorted.map((s) => s[0]) });
        }
        console.log(`  Expense Sheet: ${map.size} ledger→MAIN GROUP mappings loaded (${ambiguous.length} ambiguous)`);
    } catch (err) {
        console.warn(`  Warning: Could not read Expense Sheet: ${err.message}`);
    }
    return { map, ambiguous };
}

function normalizeCode(code) {
    return String(code || '').replace(/[^0-9]/g, '');
}

// Read Files/TALLY Company_List.xlsx ("Company List" sheet) and build
// { '26-27': [companyName, ...], '25-26': [...], '24-25': [...] } from whichever
// rows carry a 1 under that year column. Rows marked "Not To Use" / "Only When
// Asked" simply have no 1s in any code column, so they're naturally excluded
// unless the user names them explicitly elsewhere.
function loadCompanyCodeMap(file) {
    const codeMap = new Map(); // normalized code -> [company names]
    if (!fs.existsSync(file)) {
        console.log(`  Company list file not found (${file}) — skipping company-code filtering`);
        return codeMap;
    }
    try {
        const wb = XLSX.readFile(file);
        const ws = wb.Sheets['Company List'];
        if (!ws) return codeMap;
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        // header row is the first row containing 'Company Name'
        const headerRowIdx = rows.findIndex((r) => r[0] === 'Company Name');
        if (headerRowIdx === -1) return codeMap;
        const header = rows[headerRowIdx].map((h) => String(h || '').trim());
        const nameCol = header.indexOf('Company Name');
        const codeCols = []; // { idx, normCode }
        header.forEach((h, i) => {
            const norm = normalizeCode(h);
            if (norm.length === 4) codeCols.push({ idx: i, normCode: norm });
        });
        for (let r = headerRowIdx + 1; r < rows.length; r++) {
            const name = String(rows[r][nameCol] || '').trim();
            if (!name) continue;
            for (const { idx, normCode } of codeCols) {
                if (rows[r][idx] === 1) {
                    if (!codeMap.has(normCode)) codeMap.set(normCode, []);
                    codeMap.get(normCode).push(name);
                }
            }
        }
    } catch (err) {
        console.warn(`  Warning: Could not read company list file: ${err.message}`);
    }
    return codeMap;
}

// Compare the companies actually open in Tally against what the requested
// company code requires. Returns the filtered (required ∩ open) list to
// process, and prints missing/extra so the user can fix Tally before re-running.
function resolveCompaniesForCode(openCompanies, codeMap, code) {
    const norm = normalizeCode(code);
    const required = codeMap.get(norm);
    if (!required) {
        console.log(`  No company-code mapping for "${code}" found — processing all ${openCompanies.length} open companies`);
        return { toProcess: openCompanies, missing: [], extra: [] };
    }
    const openSet = new Set(openCompanies);
    const missing = required.filter((c) => !openSet.has(c));
    const requiredSet = new Set(required);
    const extra = openCompanies.filter((c) => !requiredSet.has(c));
    const toProcess = required.filter((c) => openSet.has(c));
    return { toProcess, missing, extra };
}

function classifyRow(row, partyTypeRef, expenseMainGroupRef) {
    const vtype = row['Voucher Type'] || '';
    const party = row['Party']        || '';
    const sg    = row['Sub Group']    || '';
    const date  = row['Date']         || '';

    row['Final Voucher'] =
        /PAYMENT/i.test(vtype)         ? 'Payment' :
        /RECEIPT|RECIPT/i.test(vtype)  ? 'Receipt' : '';

    row['Voucher Classification'] =
        CHQ_RETURN_GUID_OVERRIDES.has(row['Voucher GUID']) ? 'CHQ.Ret' :
        /CHQ\.R/i.test(vtype) ? 'CHQ.Ret' :
        /CASH/i.test(vtype)   ? 'Cash'    : 'Bank';

    row['Month'] = dateToMonth(date);

    row['Expense Main Group'] = expenseMainGroupRef?.get(party.trim()) || '';

    // Priority 1: hardcoded overrides (wrong ledger group in Tally)
    const override = PARTY_OVERRIDES.get(party.trim());
    if (override) {
        row['TYPE'] = override.type; row['Category'] = override.category; row['BS/PL'] = override.bspl;
        return;
    }

    // Priority 2: manually-classified reference file
    const ref = partyTypeRef?.get(party.trim());
    if (ref && ref.type) {
        row['TYPE']     = ref.type;
        row['Category'] = ref.category || (ref.type === 'Branch' ? 'Branch' : ref.type === 'Related' ? 'Related' : 'OTH');
        row['BS/PL']    = ref.bspl || TYPE_BSPL[ref.type] || '';
        return;
    }

    // Priority 3: sub-group pattern logic
    const type = getType(sg, party);
    row['TYPE']     = type;
    row['Category'] = type === 'Branch' ? 'Branch' : type === 'Related' ? 'Related' : 'OTH';
    row['BS/PL']    = TYPE_BSPL[type] || '';
}

// ---------------------------------------------------------------------------
// Tally HTTP
// ---------------------------------------------------------------------------

function postToTally(xml) {
    return new Promise((resolve, reject) => {
        const req = http.request(
            {
                hostname: TALLY_HOST,
                port: TALLY_PORT,
                path: '',
                method: 'POST',
                headers: {
                    'Content-Length': Buffer.byteLength(xml, 'utf16le'),
                    'Content-Type': 'text/xml;charset=utf-16',
                },
            },
            (res) => {
                let data = '';
                res.setEncoding('utf16le')
                    .on('data', (c) => { data += c; })
                    .on('end', () => resolve(data));
            }
        );
        req.on('error', (err) => {
            if (err.code === 'ECONNREFUSED')
                reject(new Error(
                    `Cannot connect to Tally at ${TALLY_HOST}:${TALLY_PORT}.\n` +
                    `Ensure Tally is running and XML server is enabled:\n` +
                    `  Help (F1) > Settings > Connectivity > Tally is acting as: Server\n` +
                    `If Tally is running, confirm the port it is actually listening on and\n` +
                    `update TALLY_PORT in the project .env to match.`
                ));
            else
                reject(err);
        });
        req.write(xml, 'utf16le');
        req.end();
    });
}

// ---------------------------------------------------------------------------
// TDL XML — company list + details
// ---------------------------------------------------------------------------

const XML_GET_COMPANY_DETAILS =
    `<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Data</TYPE><ID>GetCoDtlBT</ID></HEADER>` +
    `<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>` +
    `<TDL><TDLMESSAGE>` +
    `<REPORT NAME="GetCoDtlBT"><FORMS>CDFBT</FORMS></REPORT>` +
    `<FORM NAME="CDFBT"><PARTS>CDPBT</PARTS><XMLTAG>DATA</XMLTAG></FORM>` +
    `<PART NAME="CDPBT"><LINES>CDLBT</LINES><REPEAT>CDLBT : CDCBT</REPEAT><SCROLLED>Vertical</SCROLLED></PART>` +
    `<LINE NAME="CDLBT"><FIELDS>FCNBT,FCSBT</FIELDS><XMLTAG>ROW</XMLTAG></LINE>` +
    `<FIELD NAME="FCNBT"><SET>$Name</SET><XMLTAG>NAME</XMLTAG></FIELD>` +
    `<FIELD NAME="FCSBT"><SET>$StateName</SET><XMLTAG>STATE</XMLTAG></FIELD>` +
    `<COLLECTION NAME="CDCBT"><TYPE>Company</TYPE></COLLECTION>` +
    `</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>`;

const XML_GET_COMPANIES =
    `<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Data</TYPE><ID>GetCoBT</ID></HEADER>` +
    `<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>` +
    `<TDL><TDLMESSAGE>` +
    `<REPORT NAME="GetCoBT"><FORMS>CoFBT</FORMS></REPORT>` +
    `<FORM NAME="CoFBT"><PARTS>CoPBT</PARTS><XMLTAG>DATA</XMLTAG></FORM>` +
    `<PART NAME="CoPBT"><LINES>CoLBT</LINES><REPEAT>CoLBT : CoCBT</REPEAT><SCROLLED>Vertical</SCROLLED></PART>` +
    `<LINE NAME="CoLBT"><FIELDS>FNBT</FIELDS><XMLTAG>ROW</XMLTAG></LINE>` +
    `<FIELD NAME="FNBT"><SET>$Name</SET><XMLTAG>NAME</XMLTAG></FIELD>` +
    `<COLLECTION NAME="CoCBT"><TYPE>Company</TYPE></COLLECTION>` +
    `</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>`;

// ---------------------------------------------------------------------------
// TDL XML — group hierarchy
// ---------------------------------------------------------------------------

function buildGroupXml(companyName, tag) {
    const co = xmlEscape(companyName);
    return (
        `<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Data</TYPE><ID>GetGrp${tag}</ID></HEADER>` +
        `<BODY><DESC>` +
        `<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT><SVCURRENTCOMPANY>${co}</SVCURRENTCOMPANY></STATICVARIABLES>` +
        `<TDL><TDLMESSAGE>` +
        `<REPORT NAME="GetGrp${tag}"><FORMS>GF${tag}</FORMS></REPORT>` +
        `<FORM NAME="GF${tag}"><PARTS>GP${tag}</PARTS><XMLTAG>DATA</XMLTAG></FORM>` +
        `<PART NAME="GP${tag}"><LINES>GL${tag}</LINES><REPEAT>GL${tag} : GC${tag}</REPEAT><SCROLLED>Vertical</SCROLLED></PART>` +
        `<LINE NAME="GL${tag}"><FIELDS>FGN${tag},FGP${tag}</FIELDS><XMLTAG>ROW</XMLTAG></LINE>` +
        `<FIELD NAME="FGN${tag}"><SET>$Name</SET><XMLTAG>NAME</XMLTAG></FIELD>` +
        `<FIELD NAME="FGP${tag}"><SET>$Parent</SET><XMLTAG>PARENT</XMLTAG></FIELD>` +
        `<COLLECTION NAME="GC${tag}"><TYPE>Group</TYPE></COLLECTION>` +
        `</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>`
    );
}

// ---------------------------------------------------------------------------
// TDL XML — vouchers
// ---------------------------------------------------------------------------

function buildVoucherXml(companyName, tag, filterFormula) {
    const co = xmlEscape(companyName);
    let ledFields = '';
    let ledFieldNames = '';
    for (let i = 1; i <= MAX_LED; i++) {
        ledFields +=
            `<FIELD NAME="FL${i}N${tag}"><SET>$AllLedgerEntries[${i}].LedgerName</SET><XMLTAG>L${i}N</XMLTAG></FIELD>` +
            `<FIELD NAME="FL${i}P${tag}"><SET>$Parent:Ledger:$AllLedgerEntries[${i}].LedgerName</SET><XMLTAG>L${i}P</XMLTAG></FIELD>` +
            `<FIELD NAME="FL${i}A${tag}"><SET>$$StringFindAndReplace:(if $$IsDebit:$AllLedgerEntries[${i}].Amount then -$$NumValue:$AllLedgerEntries[${i}].Amount else $$NumValue:$AllLedgerEntries[${i}].Amount):"(-)":"-"</SET><XMLTAG>L${i}A</XMLTAG></FIELD>`;
        ledFieldNames += `,FL${i}N${tag},FL${i}P${tag},FL${i}A${tag}`;
    }
    return (
        `<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Data</TYPE><ID>BankRpt${tag}</ID></HEADER>` +
        `<BODY><DESC>` +
        `<STATICVARIABLES>` +
        `<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>` +
        `<SVFROMDATE>${FROM_DATE}</SVFROMDATE><SVTODATE>${TO_DATE}</SVTODATE>` +
        `<SVCURRENTCOMPANY>${co}</SVCURRENTCOMPANY>` +
        `</STATICVARIABLES>` +
        `<TDL><TDLMESSAGE>` +
        `<REPORT NAME="BankRpt${tag}"><FORMS>BForm${tag}</FORMS></REPORT>` +
        `<FORM NAME="BForm${tag}"><PARTS>BPart${tag}</PARTS><XMLTAG>DATA</XMLTAG></FORM>` +
        `<PART NAME="BPart${tag}"><LINES>BLine${tag}</LINES><REPEAT>BLine${tag} : BColl${tag}</REPEAT><SCROLLED>Vertical</SCROLLED></PART>` +
        `<LINE NAME="BLine${tag}"><FIELDS>FDt${tag},FGd${tag},FVn${tag},FVt${tag},FNr${tag},FEnt${tag},FMnt${tag},FAlt${tag}${ledFieldNames}</FIELDS><XMLTAG>ROW</XMLTAG></LINE>` +
        `<FIELD NAME="FDt${tag}"><SET>$$PyrlYYYYMMDDFormat:$Date:"-"</SET><XMLTAG>DATE</XMLTAG></FIELD>` +
        `<FIELD NAME="FGd${tag}"><SET>$GUID</SET><XMLTAG>GUID</XMLTAG></FIELD>` +
        `<FIELD NAME="FVn${tag}"><SET>$VoucherNumber</SET><XMLTAG>VCHNO</XMLTAG></FIELD>` +
        `<FIELD NAME="FVt${tag}"><SET>$VoucherTypeName</SET><XMLTAG>VCHTYPE</XMLTAG></FIELD>` +
        `<FIELD NAME="FNr${tag}"><SET>$Narration</SET><XMLTAG>NARRATION</XMLTAG></FIELD>` +
        `<FIELD NAME="FEnt${tag}"><SET>$EnteredBy</SET><XMLTAG>ENTBY</XMLTAG></FIELD>` +
        `<FIELD NAME="FMnt${tag}"><SET>$LastMntdByUser</SET><XMLTAG>MNTBY</XMLTAG></FIELD>` +
        `<FIELD NAME="FAlt${tag}"><SET>$$PyrlYYYYMMDDFormat:$AlteredDate:"-"</SET><XMLTAG>ALTDT</XMLTAG></FIELD>` +
        ledFields +
        `<COLLECTION NAME="BColl${tag}"><TYPE>Voucher</TYPE>` +
        `<FETCH>AllLedgerEntries,Narration</FETCH>` +
        `<FILTER>Filt${tag},FiltNC${tag},FiltNO${tag}</FILTER></COLLECTION>` +
        `<SYSTEM TYPE="Formulae" NAME="Filt${tag}">${filterFormula}</SYSTEM>` +
        `<SYSTEM TYPE="Formulae" NAME="FiltNC${tag}">NOT $IsCancelled</SYSTEM>` +
        `<SYSTEM TYPE="Formulae" NAME="FiltNO${tag}">NOT $IsOptional</SYSTEM>` +
        `</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>`
    );
}

// ---------------------------------------------------------------------------
// Parsers
// ---------------------------------------------------------------------------

const xmlParser = new XMLParser({
    parseTagValue: false,
    isArray: (tag) => tag === 'ROW',
});

function xmlEscape(str) {
    return str
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function unescape(str) {
    if (!str) return '';
    return str
        .replace(/&amp;/g, '&').replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>').replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'").replace(/&#\d+;/g, '');
}

function parseAmount(str) {
    if (!str) return 0;
    const n = parseFloat(String(str).replace(/[(),\s]/g, ''));
    return isNaN(n) ? 0 : n;
}

function parseTallyDate(str) {
    if (!str || str.length !== 10) return '';
    const [y, m, d] = str.split('-');
    return `${d}-${m}-${y}`;
}

function parseCompanies(xml) {
    try {
        const obj = xmlParser.parse(xml);
        if (!obj.DATA || !Array.isArray(obj.DATA.ROW)) return [];
        return obj.DATA.ROW.map((r) => unescape(r.NAME || '')).filter(Boolean);
    } catch { return []; }
}

function parseCompanyDetails(xml) {
    const map = new Map();
    try {
        const obj = xmlParser.parse(xml);
        for (const r of (obj.DATA?.ROW || [])) {
            const name  = unescape(r.NAME  || '');
            const state = unescape(r.STATE || '');
            if (name) map.set(name, { state });
        }
    } catch { /* ignore */ }
    return map;
}

function parseGroupHierarchy(xml) {
    const map = new Map();
    try {
        const obj = xmlParser.parse(xml);
        for (const r of (obj.DATA?.ROW || [])) {
            const name   = unescape(r.NAME   || '').trim();
            const parent = unescape(r.PARENT || '').trim();
            if (name) map.set(name, parent);
        }
    } catch { /* ignore */ }
    return map;
}

function getMainGroup(groupName, groupMap) {
    if (!groupName) return '';
    let current = groupName;
    const visited = new Set();
    while (true) {
        if (visited.has(current)) return current;
        visited.add(current);
        const parent = groupMap.get(current);
        if (!parent || parent.toLowerCase() === 'primary') return current;
        current = parent;
    }
}

function parseVouchers(xml, companyName, groupMap, isPayment) {
    try {
        const obj = xmlParser.parse(xml);
        if (!obj.DATA || !Array.isArray(obj.DATA.ROW)) return [];

        const results = [];
        for (const r of obj.DATA.ROW) {
            const bankEntries  = [];
            const partyEntries = [];

            for (let i = 1; i <= MAX_LED; i++) {
                const ledName = unescape(r[`L${i}N`] || '');
                if (!ledName) continue;
                const signed = parseAmount(r[`L${i}A`]);
                if (Math.abs(signed) < 0.005) continue;
                const parent = unescape(r[`L${i}P`] || '');

                if (isPayment) {
                    if (signed > 0)
                        bankEntries.push({ name: ledName, amount: signed });
                    else
                        partyEntries.push({ name: ledName, parent, amount: Math.abs(signed) });
                } else {
                    if (signed < 0) {
                        if (isLedgerInBankGroup(parent, groupMap))
                            bankEntries.push({ name: ledName, amount: Math.abs(signed), _isBankGroup: true });
                        else
                            bankEntries.push({ name: ledName, amount: Math.abs(signed), _isBankGroup: false, parent });
                    } else {
                        partyEntries.push({ name: ledName, parent, amount: signed });
                    }
                }
            }

            if (!isPayment) {
                const hasBankGroup = bankEntries.some((e) => e._isBankGroup);
                if (hasBankGroup) {
                    const realBank = bankEntries.filter((e) => e._isBankGroup);
                    const nonBank  = bankEntries.filter((e) => !e._isBankGroup);
                    bankEntries.length = 0;
                    bankEntries.push(...realBank);
                    for (const nb of nonBank)
                        partyEntries.push({ name: nb.name, parent: nb.parent || '', amount: nb.amount });
                }
            }

            if (bankEntries.length === 0) continue;

            const bankName   = bankEntries.map((e) => e.name).join('; ');
            const bankAmount = bankEntries.reduce((s, e) => s + e.amount, 0);
            const dateSort   = r.DATE || '';
            const date       = parseTallyDate(dateSort);
            const guid       = unescape(r.GUID || '');
            const vchNo      = unescape(r.VCHNO || '');
            const vchType    = unescape(r.VCHTYPE || '');
            const narration  = unescape(r.NARRATION || '');
            const entBy      = unescape(r.ENTBY  || '');
            const mntBy      = unescape(r.MNTBY  || '');
            const altDt      = r.ALTDT || '';
            // Modified On / Last Modified By — only populated if voucher was altered after entry
            const wasAltered = altDt && altDt !== dateSort;
            const modDate    = wasAltered ? parseTallyDate(altDt) : '';
            const lastMntBy  = wasAltered ? mntBy : '';

            for (const entry of partyEntries) {
                const subGroup  = entry.parent;
                const mainGroup = getMainGroup(subGroup, groupMap);
                results.push({
                    Company:              companyName,
                    _dateSort:            dateSort,
                    Date:                 date,
                    'Voucher GUID':       guid,
                    'Voucher No':         vchNo,
                    'Voucher Type':       vchType,
                    Party:                entry.name,
                    'Sub Group':          subGroup,
                    'Main Group':         mainGroup,
                    'Amount (₹)':         entry.amount,
                    'Bank Account':       bankName,
                    'Bank Amount (₹)':    bankAmount,
                    Narration:            narration,
                    'Entered By':         entBy,
                    'Last Modified By':   lastMntBy,
                    'Modified On':        modDate,
                });
            }
        }
        return results;
    } catch (e) {
        console.error('  Parse error:', e.message);
        return [];
    }
}

// ---------------------------------------------------------------------------
// Excel helpers — flat sheet
// ---------------------------------------------------------------------------

const HEADERS = [
    'Company', 'Company State', 'Date', 'Voucher GUID', 'Voucher No',
    'Voucher Type', 'Party', 'Sub Group', 'Main Group', 'Amount (₹)',
    'Bank Account', 'Bank Amount (₹)', 'Narration',
    'Final Voucher', 'Voucher Classification', 'Category', 'TYPE', 'BS/PL', 'Month',
    'Co Code', 'Location', 'Expense Main Group',
    'Entered By', 'Last Modified By', 'Modified On',
];
const COL_WIDTHS = [
    32, 16, 14, 38, 18,
    22, 40, 30, 30, 16,
    50, 16, 55,
    14, 22, 12, 18, 10, 10,
    12, 10, 26,
    20, 20, 14,
];

function makeSheet(rows, includeCompanyCol = true) {
    const cols   = includeCompanyCol ? HEADERS : HEADERS.slice(2);
    const widths = includeCompanyCol ? COL_WIDTHS : COL_WIDTHS.slice(2);
    const data   = rows.map((r) => cols.map((h) => r[h] ?? ''));
    const ws     = XLSX.utils.aoa_to_sheet([cols, ...data]);
    ws['!cols'] = widths.map((w) => ({ wch: w }));
    ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: cols.length - 1 } }),
    };
    return ws;
}

function makeSummarySheet(companies, companyInfoMap, payments, receipts) {
    const now = new Date();
    const refreshedAt = now.toLocaleString('en-IN', {
        day: '2-digit', month: 'short', year: 'numeric',
        hour: '2-digit', minute: '2-digit', hour12: true,
    });
    const rows = [
        [`Last Refreshed: ${refreshedAt}`],
        [],
        ['Company', 'Company State', 'Payment Vouchers', 'Payment Amount (₹)', 'Receipt Vouchers', 'Receipt Amount (₹)'],
    ];
    let gPayCount = 0, gPayAmt = 0, gRecCount = 0, gRecAmt = 0;
    for (const co of companies) {
        const info  = companyInfoMap.get(co) || { state: '' };
        const pRows = payments.filter((v) => v.Company === co);
        const rRows = receipts.filter((v) => v.Company === co);
        let payCount = 0, payAmt = 0, recCount = 0, recAmt = 0;
        const seenP = new Set(), seenR = new Set();
        for (const v of pRows) {
            if (!seenP.has(v['Voucher GUID'])) { seenP.add(v['Voucher GUID']); payCount++; payAmt += v['Bank Amount (₹)']; }
        }
        for (const v of rRows) {
            if (!seenR.has(v['Voucher GUID'])) { seenR.add(v['Voucher GUID']); recCount++; recAmt += v['Bank Amount (₹)']; }
        }
        gPayCount += payCount; gPayAmt += payAmt;
        gRecCount += recCount; gRecAmt += recAmt;
        rows.push([co, info.state, payCount, Math.round(payAmt * 100) / 100, recCount, Math.round(recAmt * 100) / 100]);
    }
    rows.push(['TOTAL', '', gPayCount, Math.round(gPayAmt * 100) / 100, gRecCount, Math.round(gRecAmt * 100) / 100]);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 45 }, { wch: 16 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 22 }];
    return ws;
}

// ---------------------------------------------------------------------------
// Pivot sheet — ALL / BANK / CASH / RETUNN
// ---------------------------------------------------------------------------

function r2(n) { return Math.round(n * 100) / 100; }

function makePivotSheet(allRows, filterFn, months) {
    const rows = filterFn ? allRows.filter(filterFn) : allRows;

    // Aggregate [fv][bspl][type][month] = sum
    const agg = {};
    for (const row of rows) {
        const fv   = row['Final Voucher']           || '';
        const bspl = row['BS/PL']                   || '';
        const type = row['TYPE']                    || '';
        const mon  = row['Month']                   || '';
        const amt  = Number(row['Amount (₹)'])      || 0;
        if (!agg[fv])             agg[fv] = {};
        if (!agg[fv][bspl])       agg[fv][bspl] = {};
        if (!agg[fv][bspl][type]) agg[fv][bspl][type] = {};
        agg[fv][bspl][type][mon] = (agg[fv][bspl][type][mon] || 0) + amt;
    }

    const FV_ORDER   = ['Payment', 'Receipt'];
    const BSPL_ORDER = ['BS', 'PL', 'Loan'];

    const out = [['Final Voucher', 'BS/PL', 'TYPE', ...months, 'Grand Total']];
    const grandByMon = {};

    const fvKeys = [
        ...FV_ORDER.filter((k) => agg[k]),
        ...Object.keys(agg).filter((k) => !FV_ORDER.includes(k) && agg[k]),
    ];

    for (const fv of fvKeys) {
        const fvByMon = {};
        let firstFvRow = true;

        const bsplKeys = [
            ...BSPL_ORDER.filter((k) => agg[fv][k]),
            ...Object.keys(agg[fv]).filter((k) => !BSPL_ORDER.includes(k)),
        ];

        for (const bspl of bsplKeys) {
            const bsplByMon = {};
            let firstBsplRow = true;

            const typeKeys = Object.keys(agg[fv][bspl]).sort((a, b) => {
                const totA = Object.values(agg[fv][bspl][a]).reduce((s, v) => s + v, 0);
                const totB = Object.values(agg[fv][bspl][b]).reduce((s, v) => s + v, 0);
                return totB - totA;
            });
            for (const type of typeKeys) {
                const byMon = agg[fv][bspl][type];
                let rowGT = 0;
                const rowVals = months.map((m) => {
                    const v = r2(byMon[m] || 0);
                    rowGT        += v;
                    bsplByMon[m]  = (bsplByMon[m]  || 0) + v;
                    fvByMon[m]    = (fvByMon[m]    || 0) + v;
                    grandByMon[m] = (grandByMon[m] || 0) + v;
                    return v;
                });
                out.push([firstFvRow && firstBsplRow ? fv : '',
                          firstBsplRow ? bspl : '',
                          type, ...rowVals, r2(rowGT)]);
                firstFvRow = firstBsplRow = false;
            }

            // BSPL subtotal
            let bsplGT = 0;
            const bsplVals = months.map((m) => { const v = r2(bsplByMon[m] || 0); bsplGT += v; return v; });
            out.push(['', `${bspl} Total`, '', ...bsplVals, r2(bsplGT)]);
        }

        // FV total
        let fvGT = 0;
        const fvVals = months.map((m) => { const v = r2(fvByMon[m] || 0); fvGT += v; return v; });
        out.push([`${fv} Total`, '', '', ...fvVals, r2(fvGT)]);
    }

    // Grand total
    let overallGT = 0;
    const grandVals = months.map((m) => { const v = r2(grandByMon[m] || 0); overallGT += v; return v; });
    out.push(['Grand Total', '', '', ...grandVals, r2(overallGT)]);

    const ws = XLSX.utils.aoa_to_sheet(out);
    ws['!cols'] = [
        { wch: 16 }, { wch: 12 }, { wch: 22 },
        ...months.map(() => ({ wch: 14 })),
        { wch: 16 },
    ];
    return ws;
}

// ---------------------------------------------------------------------------
// Party sheet — Related Partry / Bracnch Mis
// ---------------------------------------------------------------------------

function makePartySheet(allRows, categoryFilter) {
    const rows = allRows.filter((r) => r['Category'] === categoryFilter);

    // Determine column combos that actually exist
    const COL_ORDER = [
        'Bank Payment', 'Bank Receipt',
        'Cash Payment', 'Cash Receipt',
        'CHQ.Ret Payment', 'CHQ.Ret Receipt',
    ];
    const colSet = new Set();
    for (const r of rows) {
        const vc = r['Voucher Classification'] || '';
        const fv = r['Final Voucher']          || '';
        if (vc && fv) colSet.add(`${vc} ${fv}`);
    }
    const sortedCols = [
        ...COL_ORDER.filter((c) => colSet.has(c)),
        ...[...colSet].filter((c) => !COL_ORDER.includes(c)).sort(),
    ];

    // Aggregate party → colKey → sum
    const partyMap = new Map();
    for (const r of rows) {
        const party  = r['Party']                 || '';
        const vc     = r['Voucher Classification'] || '';
        const fv     = r['Final Voucher']          || '';
        const amt    = Number(r['Amount (₹)'])     || 0;
        const colKey = `${vc} ${fv}`;
        if (!partyMap.has(party)) partyMap.set(party, {});
        const m = partyMap.get(party);
        m[colKey] = (m[colKey] || 0) + amt;
    }

    const header = ['Party', ...sortedCols, 'Grand Total'];
    const out    = [header];
    const grandTotals = {};

    for (const [party, colTotals] of [...partyMap.entries()].sort((a, b) => {
        const totA = Object.values(a[1]).reduce((s, v) => s + v, 0);
        const totB = Object.values(b[1]).reduce((s, v) => s + v, 0);
        return totB - totA;
    })) {
        let rowGT = 0;
        const vals = sortedCols.map((c) => {
            const v = r2(colTotals[c] || 0);
            rowGT        += v;
            grandTotals[c] = (grandTotals[c] || 0) + v;
            return v;
        });
        out.push([party, ...vals, r2(rowGT)]);
    }

    // Grand total row
    let overallGT = 0;
    const gtVals = sortedCols.map((c) => { const v = r2(grandTotals[c] || 0); overallGT += v; return v; });
    out.push(['Grand Total', ...gtVals, r2(overallGT)]);

    const ws = XLSX.utils.aoa_to_sheet(out);
    ws['!cols'] = [
        { wch: 48 },
        ...sortedCols.map(() => ({ wch: 16 })),
        { wch: 16 },
    ];
    ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: out.length - 1, c: header.length - 1 } }),
    };
    return ws;
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function safeName(name, usedNames) {
    const base = name.replace(/[\\\/\?\*\[\]:]/g, '-').substring(0, 31);
    if (!usedNames.has(base)) { usedNames.add(base); return base; }
    for (let i = 2; i < 100; i++) {
        const suffix = ` (${i})`;
        const candidate = base.substring(0, 31 - suffix.length) + suffix;
        if (!usedNames.has(candidate)) { usedNames.add(candidate); return candidate; }
    }
    return base;
}

function sortByDate(rows) {
    return rows.slice().sort((a, b) => {
        const dc = (a._dateSort || '').localeCompare(b._dateSort || '');
        return dc !== 0 ? dc : (a['Voucher GUID'] || '').localeCompare(b['Voucher GUID'] || '');
    });
}

function extractMonths(rows) {
    const seen = new Set();
    for (const r of rows) { if (r['Month']) seen.add(r['Month']); }
    return [...seen].sort((a, b) => {
        const [ma, ya] = a.split('-');
        const [mb, yb] = b.split('-');
        if (ya !== yb) return parseInt(ya) - parseInt(yb);
        return MONTH_ORDER.indexOf(ma) - MONTH_ORDER.indexOf(mb);
    });
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
    console.log('=== Tally Bank Transactions Fetcher (Payments + Receipts) ===');
    console.log(`Period  : ${FROM_DATE} to ${TO_DATE}`);
    console.log(`Tally   : ${TALLY_HOST}:${TALLY_PORT}`);
    console.log(`Output  : ${OUTPUT_FILE}\n`);

    fs.mkdirSync(OUTPUT_DIR, { recursive: true });

    console.log('Connecting to Tally and fetching company list...');
    let companiesXml, detailsXml;
    try {
        [companiesXml, detailsXml] = await Promise.all([
            postToTally(XML_GET_COMPANIES),
            postToTally(XML_GET_COMPANY_DETAILS),
        ]);
    } catch (err) {
        console.error('\nError:', err.message);
        process.exit(1);
    }

    let companies        = parseCompanies(companiesXml);
    const companyInfoMap = parseCompanyDetails(detailsXml);

    console.log('\nLoading party→TYPE reference file...');
    const partyTypeRef = await loadPartyTypeReference(REFERENCE_FILE);

    console.log('Loading Expense Sheet (ledger→MAIN GROUP cross-check)...');
    const { map: expenseMainGroupRef, ambiguous: expenseAmbiguous } = loadExpenseMainGroupRef(EXPENSE_FILE);
    if (expenseAmbiguous.length) {
        console.log(`  Ambiguous ledgers (used most-frequent MAIN GROUP): ${expenseAmbiguous.map((a) => a.ledger).join(', ')}`);
    }

    if (companies.length === 0) {
        console.error('No companies found. Ensure at least one company is open in Tally.');
        process.exit(1);
    }

    console.log(`\nCompany code: ${COMPANY_CODE}  (Files/TALLY Company_List.xlsx)`);
    const companyCodeMap = loadCompanyCodeMap(COMPANY_LIST_FILE);
    const { toProcess, missing, extra } = resolveCompaniesForCode(companies, companyCodeMap, COMPANY_CODE);
    if (missing.length > 0) {
        console.error(`\nERROR: Company code "${COMPANY_CODE}" requires companies that are NOT currently open in Tally:`);
        for (const c of missing) console.error(`  ✗ missing (please OPEN in Tally): ${c}`);
        if (extra.length > 0) {
            console.error(`\n  Currently open but not needed for this code (you can CLOSE these):`);
            for (const c of extra) console.error(`  - ${c}`);
        }
        console.error('\nOpen the missing companies in Tally (and optionally close the extras above), then re-run.');
        process.exit(1);
    }
    if (extra.length > 0) {
        console.log(`  Note: ${extra.length} open company(ies) are not part of code "${COMPANY_CODE}" and will be skipped: ${extra.join(', ')}`);
    }
    companies = toProcess;

    console.log(`Connection: OK  |  Companies (${companies.length}):`);
    for (const co of companies) {
        const info = companyInfoMap.get(co) || { state: '' };
        console.log(`  • ${co}${info.state ? '  [' + info.state + ']' : ''}`);
    }
    console.log();

    const allPayments = [];
    const allReceipts = [];

    for (const company of companies) {
        const info = companyInfoMap.get(company) || { state: '' };
        console.log(`\n[${company}]`);

        process.stdout.write(`  Groups   ... `);
        let groupMap = new Map();
        try {
            const gxml = await postToTally(buildGroupXml(company, 'BT'));
            groupMap   = parseGroupHierarchy(gxml);
            console.log(`${groupMap.size} groups`);
        } catch (err) {
            console.log(`ERROR — ${err.message}`);
        }

        process.stdout.write(`  Payments ... `);
        try {
            const xml  = await postToTally(buildVoucherXml(company, 'BP', '$$IsPayment:$VoucherTypeName'));
            const rows = parseVouchers(xml, company, groupMap, true);
            const meta = COMPANY_META.get(company) || { coCode: '', location: '' };
            for (const v of rows) { v['Company State'] = info.state; v['Co Code'] = meta.coCode; v['Location'] = meta.location; classifyRow(v, partyTypeRef, expenseMainGroupRef); }
            const seen = new Set();
            let cnt = 0, amt = 0;
            for (const v of rows) {
                if (!seen.has(v['Voucher GUID'])) { seen.add(v['Voucher GUID']); cnt++; amt += v['Bank Amount (₹)']; }
            }
            allPayments.push(...rows);
            console.log(`${cnt} vouchers → ${rows.length} entries  |  ₹${amt.toLocaleString('en-IN', { maximumFractionDigits: 2 })}`);
        } catch (err) {
            console.log(`ERROR — ${err.message}`);
        }

        process.stdout.write(`  Receipts ... `);
        try {
            const xml  = await postToTally(buildVoucherXml(company, 'BR', '$$IsReceipt:$VoucherTypeName'));
            const rows = parseVouchers(xml, company, groupMap, false);
            const metaR = COMPANY_META.get(company) || { coCode: '', location: '' };
            for (const v of rows) { v['Company State'] = info.state; v['Co Code'] = metaR.coCode; v['Location'] = metaR.location; classifyRow(v, partyTypeRef, expenseMainGroupRef); }
            const seen = new Set();
            let cnt = 0, amt = 0;
            for (const v of rows) {
                if (!seen.has(v['Voucher GUID'])) { seen.add(v['Voucher GUID']); cnt++; amt += v['Bank Amount (₹)']; }
            }
            allReceipts.push(...rows);
            console.log(`${cnt} vouchers → ${rows.length} entries  |  ₹${amt.toLocaleString('en-IN', { maximumFractionDigits: 2 })}`);
        } catch (err) {
            console.log(`ERROR — ${err.message}`);
        }
    }

    const allTransactions = sortByDate([...allPayments, ...allReceipts]);
    const months          = extractMonths(allTransactions);

    console.log(`\nMonths detected: ${months.join(', ')}`);

    // Build workbook — sheet order:
    // Summary | ALL | BANK | CASH | RETUNN | Related Partry | Bracnch Mis | All Transactions | [company sheets]
    const wb = XLSX.utils.book_new();
    const usedNames = new Set();
    const addSheet = (ws, name) => XLSX.utils.book_append_sheet(wb, ws, safeName(name, usedNames));

    addSheet(makeSummarySheet(companies, companyInfoMap, allPayments, allReceipts), 'Summary');
    addSheet(makePivotSheet(allTransactions, null, months), 'ALL');
    addSheet(makePivotSheet(allTransactions, (r) => r['Voucher Classification'] === 'Bank',    months), 'BANK');
    addSheet(makePivotSheet(allTransactions, (r) => r['Voucher Classification'] === 'Cash',    months), 'CASH');
    addSheet(makePivotSheet(allTransactions, (r) => r['Voucher Classification'] === 'CHQ.Ret', months), 'RETUNN');
    addSheet(makePartySheet(allTransactions, 'Related'), 'Related Partry');
    addSheet(makePartySheet(allTransactions, 'Branch'),  'Bracnch Mis');
    addSheet(makeSheet(allTransactions, true), 'All Transactions');

    for (const company of companies) {
        const pRows   = allPayments.filter((v) => v.Company === company);
        const rRows   = allReceipts.filter((v) => v.Company === company);
        const combined = sortByDate([...pRows, ...rRows]);
        if (combined.length === 0) continue;
        addSheet(makeSheet(combined, false), company);
    }

    XLSX.writeFile(wb, OUTPUT_FILE);

    console.log(`\n\nExcel saved: ${OUTPUT_FILE}`);
    const pSeen = new Set(), rSeen = new Set();
    let pCnt = 0, pAmt = 0, rCnt = 0, rAmt = 0;
    for (const v of allPayments) { if (!pSeen.has(v['Voucher GUID'])) { pSeen.add(v['Voucher GUID']); pCnt++; pAmt += v['Bank Amount (₹)']; } }
    for (const v of allReceipts) { if (!rSeen.has(v['Voucher GUID'])) { rSeen.add(v['Voucher GUID']); rCnt++; rAmt += v['Bank Amount (₹)']; } }
    console.log(`Payments : ${pCnt} vouchers → ${allPayments.length} entries  |  ₹${pAmt.toLocaleString('en-IN', { maximumFractionDigits: 2 })}`);
    console.log(`Receipts : ${rCnt} vouchers → ${allReceipts.length} entries  |  ₹${rAmt.toLocaleString('en-IN', { maximumFractionDigits: 2 })}`);
    console.log(`Combined : ${allTransactions.length} total entries`);

    // Build native-PivotTable report
    const pivotScript = path.join(__dirname, 'create_pivot_report.py');
    if (fs.existsSync(pivotScript)) {
        console.log('\n--- Building Pivot Report ---');
        try {
            execFileSync('python', [pivotScript], { stdio: 'inherit' });
        } catch (err) {
            console.error('Pivot report failed:', err.message);
        }
    }
}

main().catch((err) => {
    console.error('Fatal:', err);
    process.exit(1);
});
