const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cookieParser = require('cookie-parser');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'mypassword';

// --- Step 1: Global crash logging ---
process.on('uncaughtException', (err) => {
    console.error('Uncaught Exception:', err);
});
process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

// --- Step 3: Robust folder & JSON creation ---
['data', 'uploads', 'covers', 'public'].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});
if (!fs.existsSync('data/books.json')) fs.writeFileSync('data/books.json', '[]');
if (!fs.existsSync('data/votes.json')) fs.writeFileSync('data/votes.json', '{}');

// --- Middleware ---
app.use(express.static('public'));
app.use(cookieParser());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.set('view engine', 'ejs');

// --- Multer setup for Excel file ---
const upload = multer({ dest: 'uploads/' });

// --- Vote version (Option 1: automatic cookie reset) ---
let voteVersion = 1;

// --- Helpers ---
function loadBooks() {
    try {
        return fs.existsSync('data/books.json') ? JSON.parse(fs.readFileSync('data/books.json')) : [];
    } catch (err) {
        console.error('Error loading books.json, resetting to empty array', err);
        fs.writeFileSync('data/books.json', '[]');
        return [];
    }
}
function saveBooks(books) {
    fs.writeFileSync('data/books.json', JSON.stringify(books, null, 2));
}
function loadVotes() {
    try {
        return fs.existsSync('data/votes.json') ? JSON.parse(fs.readFileSync('data/votes.json')) : {};
    } catch (err) {
        console.error('Error loading votes.json, resetting to empty object', err);
        fs.writeFileSync('data/votes.json', '{}');
        return {};
    }
}
function saveVotes(votes) {
    fs.writeFileSync('data/votes.json', JSON.stringify(votes, null, 2));
}

// --- Admin routes ---
app.get('/admin', (req, res) => {
    res.render('admin', { message: '' });
});

app.post('/admin', upload.single('excel'), (req, res) => {
    const { password } = req.body;
    if (password !== ADMIN_PASSWORD) {
        return res.render('admin', { message: '❌ Wrong password!' });
    }
    if (!req.file) {
        return res.render('admin', { message: '⚠️ No file uploaded.' });
    }

    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);

        // Safe Excel mapping
        const books = rows.map(row => ({
            title: row['Title'] || 'Untitled',
            author: row['Author'] || 'Unknown',
            cover: row['Cover'] || ''
        }));

        saveBooks(books);
        fs.unlinkSync(req.file.path);
        res.render('admin', { message: '✅ Books imported successfully!' });
    } catch (err) {
        console.error('Import error:', err);
        res.render('admin', { message: '❌ Error reading Excel file.' });
    }
});

// Admin view results
app.get('/admin/results', (req, res) => {
    const votes = loadVotes();
    const books = loadBooks();

    const points = {};
    books.forEach(b => points[b.title] = 0);

    Object.values(votes).forEach(v => {
        Object.entries(v).forEach(([book, pts]) => {
            points[book] += parseInt(pts) || 0;
        });
    });

    res.render('results', { points });
});

// Reset votes (admin)
app.get('/admin/reset', (req, res) => {
    fs.writeFileSync('data/votes.json', JSON.stringify({}, null, 2));
    voteVersion++; // Increment vote version to invalidate old cookies
    res.render('admin', { message: '✅ All votes reset! Cookies are now cleared for voting.' });
});

// --- Voting routes ---
app.get('/', (req, res) => {
    const votedCookie = req.cookies.voted;
    if (votedCookie && votedCookie.endsWith(`-${voteVersion}`)) {
        return res.send('<h2>You have already voted. Thank you!</h2>');
    }

    const books = loadBooks();
    res.render('vote', { books });
});

app.post('/vote', (req, res) => {
    const votedCookie = req.cookies.voted;
    if (votedCookie && votedCookie.endsWith(`-${voteVersion}`)) {
        return res.send('<h2>You have already voted. Thank you!</h2>');
    }

    const votes = loadVotes();
    const voterId = Math.random().toString(36).substring(2, 15);
    const voteData = req.body;

    votes[voterId] = voteData;
    saveVotes(votes);

    // Set cookie with current vote version
    res.cookie('voted', `${voterId}-${voteVersion}`, { maxAge: 1000 * 60 * 60 * 24 * 365 });
    res.send('<h2>Thank you for voting!</h2>');
});

// --- Start server ---
app.listen(PORT, '0.0.0.0', () => console.log(`Server running on port ${PORT}`));
