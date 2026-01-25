const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'mypassword';

// ------------------- FOLDERS -------------------
['data', 'public'].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});
if (!fs.existsSync('data/books.json')) fs.writeFileSync('data/books.json', '[]');
if (!fs.existsSync('data/votes.json')) fs.writeFileSync('data/votes.json', '{}');

// ------------------- MIDDLEWARE -------------------
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.set('view engine', 'ejs');

// ------------------- HELPERS -------------------
function loadBooks() {
    return JSON.parse(fs.readFileSync('data/books.json'));
}

function saveBooks(books) {
    fs.writeFileSync('data/books.json', JSON.stringify(books, null, 2));
}

function loadVotes() {
    return JSON.parse(fs.readFileSync('data/votes.json'));
}

function saveVotes(votes) {
    fs.writeFileSync('data/votes.json', JSON.stringify(votes, null, 2));
}

// ------------------- BACKUP HELPER -------------------
function saveVoteBackup(username, voteData, month) {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupFile = `data/votes_backup_${month}_${timestamp}_${username}.json`;
    fs.writeFileSync(backupFile, JSON.stringify({ [username]: voteData }, null, 2));
}

// ------------------- MONTH UTILS -------------------
function getCurrentMonth() {
    const d = new Date();
    return `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}`;
}

// ------------------- RESTORE AT STARTUP -------------------
function restoreVotes() {
    const votes = {};
    const files = fs.readdirSync('data');
    files.forEach(file => {
        if (file.startsWith('votes_backup_') && file.endsWith('.json')) {
            const vote = JSON.parse(fs.readFileSync(`data/${file}`));
            Object.assign(votes, vote);
        }
    });
    saveVotes(votes);
    console.log('Votes restored from backups at startup.');
}

// ------------------- IMPORT EXCEL -------------------
function importExcelIfPresent() {
    const repoPaths = ['books.xlsx', 'data/books.xlsx'];
    for (const p of repoPaths) {
        if (fs.existsSync(p)) {
            const workbook = XLSX.readFile(p);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet);
            const books = rows.map(row => ({
                title: row['Title'] || 'Untitled',
                author: row['Author'] || 'Unknown',
                cover: row['Cover'] || ''
            }));
            saveBooks(books);
            console.log(`Imported ${books.length} books from ${p}`);
            return;
        }
    }
    console.log('No Excel file found in repo.');
}

// ------------------- STARTUP -------------------
importExcelIfPresent();
restoreVotes();

// ------------------- ROUTES -------------------

// health
app.get('/health', (req, res) => res.send('OK'));

// voting page
app.get('/', (req, res) => {
    const books = loadBooks();
    const currentMonth = getCurrentMonth();
    res.render('vote', { books, currentMonth });
});

// receive vote
app.post('/vote', (req, res) => {
    const { voterName, ...voteData } = req.body;
    const username = voterName?.trim();
    if (!username) return res.send('<h2>Please enter a valid name!</h2>');

    const votes = loadVotes();
    const month = getCurrentMonth();

    // Initialize month object if needed
    if (!votes[month]) votes[month] = {};

    // Check uniqueness for this month
    if (votes[month][username]) {
        return res.send('<h2>This name has already voted this month. Please choose another name.</h2>');
    }

    // SERVER VALIDATION: check 3-2-1 rule
    const values = Object.values(voteData).filter(v => v !== '');
    const counts = { '3': 0, '2': 0, '1': 0 };
    values.forEach(v => counts[v]++);
    if (counts['3'] !== 1 || counts['2'] !== 1 || counts['1'] !== 1) {
        return res.send('<h2>Invalid vote. You must assign 3, 2, and 1 exactly once.</h2>');
    }

    // Save vote
    votes[month][username] = {
        time: new Date().toISOString(),
        votes: voteData
    };
    saveVotes(votes);

    // Backup
    saveVoteBackup(username, votes[month][username], month);

    res.send('<h2>Thank you for voting!</h2>');
});

// admin page
app.get('/admin', (req, res) => {
    res.render('admin', { message: '' });
});

// results (password protected)
app.get('/admin/results', (req, res) => {
    const { password, month } = req.query;
    if (password !== ADMIN_PASSWORD) return res.send('<h2>Unauthorized</h2>');

    const votes = loadVotes();
    const books = loadBooks();
    const targetMonth = month || getCurrentMonth();

    if (!votes[targetMonth]) return res.send('<h2>No votes for this month yet.</h2><a href="/admin">Back</a>');

    const points = {};
    books.forEach(b => points[b.title] = 0);
    Object.values(votes[targetMonth]).forEach(v => {
        Object.entries(v.votes).forEach(([book, pts]) => {
            points[book] += parseInt(pts) || 0;
        });
    });

    const sorted = Object.entries(points).sort((a, b) => b[1] - a[1]);
    res.render('results', { sorted, votes: votes[targetMonth] });
});

// reset votes (password protected)
app.post('/admin/reset', (req, res) => {
    const { password } = req.body;
    if (password !== ADMIN_PASSWORD) return res.send('<h2>Wrong password</h2>');

    const votes = loadVotes();
    const month = getCurrentMonth();
    votes[month] = {}; // clear current month
    saveVotes(votes);

    // Delete backups for current month
    const files = fs.readdirSync('data');
    files.forEach(file => {
        if (file.startsWith(`votes_backup_${month}_`) && file.endsWith('.json')) {
            fs.unlinkSync(path.join('data', file));
        }
    });

    res.send('<h2>Votes reset for this month.</h2><a href="/admin">Back</a>');
});

// ------------------- START SERVER -------------------
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});
