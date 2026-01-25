const express = require('express');
const XLSX = require('xlsx');
const cookieParser = require('cookie-parser');
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
app.use(cookieParser());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.set('view engine', 'ejs');

// ------------------- COOKIE VERSION -------------------
let voteVersion = 1;

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

// ------------------- IMPORT EXCEL FROM REPO -------------------
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

// run on startup
importExcelIfPresent();

// ------------------- ROUTES -------------------

// health
app.get('/health', (req, res) => res.send('OK'));

// voting page
app.get('/', (req, res) => {
    const votedCookie = req.cookies.voted;
    if (votedCookie && votedCookie.endsWith(`-${voteVersion}`)) {
        return res.send('<h2>You have already voted. Thank you!</h2>');
    }

    const books = loadBooks();
    res.render('vote', { books });
});

// receive vote
app.post('/vote', (req, res) => {
    const votedCookie = req.cookies.voted;
    if (votedCookie && votedCookie.endsWith(`-${voteVersion}`)) {
        return res.send('<h2>You have already voted. Thank you!</h2>');
    }

    const voteData = req.body;

    // SERVER VALIDATION
    const values = Object.values(voteData).filter(v => v !== '');
    const counts = { '3': 0, '2': 0, '1': 0 };
    values.forEach(v => counts[v]++);

    if (counts['3'] !== 1 || counts['2'] !== 1 || counts['1'] !== 1) {
        return res.send('<h2>Invalid vote. You must assign 3, 2, and 1 exactly once.</h2>');
    }

    const votes = loadVotes();
    const voterId = Math.random().toString(36).substring(2, 15);

    votes[voterId] = {
        time: new Date().toISOString(),
        votes: voteData
    };

    saveVotes(votes);

    res.cookie('voted', `${voterId}-${voteVersion}`, {
        maxAge: 1000 * 60 * 60 * 24 * 365
    });

    res.send('<h2>Thank you for voting!</h2>');
});

// admin page
app.get('/admin', (req, res) => {
    res.render('admin', { message: '' });
});

// results (password protected)
app.get('/admin/results', (req, res) => {
    const { password } = req.query;
    if (password !== ADMIN_PASSWORD) {
        return res.send('<h2>Unauthorized</h2>');
    }

    const votes = loadVotes();
    const books = loadBooks();

    const points = {};
    books.forEach(b => points[b.title] = 0);

    Object.values(votes).forEach(v => {
        Object.entries(v.votes).forEach(([book, pts]) => {
            points[book] += parseInt(pts) || 0;
        });
    });

    const sorted = Object.entries(points)
        .sort((a, b) => b[1] - a[1]);

    res.render('results', { sorted, votes });
});

// reset votes (password protected)
app.post('/admin/reset', (req, res) => {
    const { password } = req.body;

    if (password !== ADMIN_PASSWORD) {
        return res.send('<h2>Wrong password</h2>');
    }

    fs.writeFileSync('data/votes.json', JSON.stringify({}, null, 2));
    voteVersion++;

    res.send('<h2>Votes reset.</h2><a href="/admin">Back</a>');
});

// ------------------- START SERVER -------------------
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});
