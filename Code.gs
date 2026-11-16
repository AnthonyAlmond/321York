/**
 * Serves the interactive House Artwork Gallery & Trivia web app.
 */
function doGet() {
  let html = '';
  try {
    const data = getAppData();
    html = buildGalleryHtml_(data);
  } catch (err) {
    html = '<div style="color:#f3efea;background:#141413;padding:3rem;font-family:sans-serif;">' +
           '<h2>Error loading spreadsheet</h2><p>' + err.message + '</p></div>';
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle('321 York Art & Sculpture Collection')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Reads both the artwork sheet (Sheet1) and the Trivia tab.
 */
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Please open this script from inside your Google Sheet via Extensions -> Apps Script.');
  }

  // 1. Read Artwork Data from Sheet1
  const artSheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  const artRange = artSheet.getDataRange();
  const displayValues = artRange.getDisplayValues();
  const formulas = artRange.getFormulas();

  const rooms = [];
  const artworks = [];

  // Start at row index 1 to skip the header row (Room | Image URL | Title | Artist | Date | Description)
  for (let i = 1; i < displayValues.length; i++) {
    const row = displayValues[i];
    const room = (row[0] || '').trim();
    const rawUrl = (row[1] || '').trim();
    const formulaVal = (formulas[i][1] || '').trim();
    const title = (row[2] || '').trim();
    const artist = (row[3] || '').trim();
    const datePainted = (row[4] || '').trim();
    const description = (row[5] || '').trim();

    if (!room && !title) continue;

    const roomName = room || 'Unassigned';
    if (rooms.indexOf(roomName) === -1) {
      rooms.push(roomName);
    }

    const parsedImage = parseColumnBUrl_(rawUrl, formulaVal);

    artworks.push({
      index: artworks.length,
      room: roomName,
      imageUrl: parsedImage.directUrl,
      altDriveUrl: parsedImage.altDriveUrl,
      driveFileId: parsedImage.driveFileId,
      title: title || 'Untitled',
      artist: artist || 'Unknown Artist',
      datePainted: datePainted || 'Date Unknown',
      description: description || ''
    });
  }

  // 2. Read Trivia Data from "Trivia" tab
  // Columns: A = Question | B = Correct Answer | C = Decoy 1 | D = Decoy 2 | E = Decoy 3
  const triviaSheet = ss.getSheetByName('Trivia');
  const trivia = [];

  if (triviaSheet) {
    const triviaValues = triviaSheet.getDataRange().getDisplayValues();
    for (let j = 1; j < triviaValues.length; j++) {
      const tRow = triviaValues[j];
      const question = (tRow[0] || '').trim();
      const correctAnswer = (tRow[1] || '').trim();
      const decoy1 = (tRow[2] || '').trim();
      const decoy2 = (tRow[3] || '').trim();
      const decoy3 = (tRow[4] || '').trim();

      if (!question || !correctAnswer) continue;

      const choices = [correctAnswer];
      if (decoy1) choices.push(decoy1);
      if (decoy2) choices.push(decoy2);
      if (decoy3) choices.push(decoy3);

      trivia.push({
        id: j - 1,
        question: question,
        correctAnswer: correctAnswer,
        choices: choices
      });
    }
  }

  return { rooms: rooms, artworks: artworks, trivia: trivia };
}

/**
 * Converts Column B links (including Google Drive /file/d/.../view links) into direct image URLs.
 */
function parseColumnBUrl_(rawUrl, formulaVal) {
  let url = rawUrl;

  if (!url && formulaVal) {
    const match = formulaVal.match(/["'](https?:\/\/[^"']+)["']/i);
    if (match && match[1]) url = match[1].trim();
  }

  if (!url) {
    return { directUrl: '', altDriveUrl: '', driveFileId: '' };
  }

  if (url.indexOf('drive.google.com') !== -1) {
    const filePathMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
    const idParamMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    const driveId = (filePathMatch && filePathMatch[1]) || (idParamMatch && idParamMatch[1]) || '';

    if (driveId) {
      return {
        directUrl: 'https://lh3.googleusercontent.com/d/' + driveId + '=w1600',
        altDriveUrl: 'https://drive.google.com/thumbnail?id=' + driveId + '&sz=w1600',
        driveFileId: driveId
      };
    }
  }

  return {
    directUrl: url,
    altDriveUrl: '',
    driveFileId: ''
  };
}

/**
 * Fallback helper: If a Google Drive image file is restricted/private and blocked by the browser,
 * fetches the image directly from Drive on the server and returns a Base64 data URI.
 */
function fetchDriveImageBase64(driveFileId) {
  if (!driveFileId) return '';
  try {
    const file = DriveApp.getFileById(driveFileId);
    const blob = file.getBlob();
    const contentType = blob.getContentType() || 'image/jpeg';
    const b64 = Utilities.base64Encode(blob.getBytes());
    return 'data:' + contentType + ';base64,' + b64;
  } catch (e) {
    return '';
  }
}

/**
 * Call this function once in the Apps Script editor (Run -> authorizeDriveAccess)
 * if your Google Drive images are private and need DriveApp read permission.
 */
function authorizeDriveAccess() {
  DriveApp.getStorageUsed();
  SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Builds the complete HTML page with Base64-encoded JSON payload.
 */
function buildGalleryHtml_(data) {
  const safeBase64Json = Utilities.base64Encode(
    Utilities.newBlob(JSON.stringify(data), 'application/json').getBytes()
  );

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>321 York Art & Sculpture Collection</title>
  <style>
    :root {
      --bg: #141413; --card-bg: #1f1e1b; --accent: #c9a66b;
      --text: #f3efea; --muted: #a8a29e; --border: #33302b;
      --success: #4ade80; --success-bg: rgba(74, 222, 128, 0.12);
      --danger: #f87171; --danger-bg: rgba(248, 113, 113, 0.12);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Georgia, serif; background: var(--bg); color: var(--text); min-height: 100vh; }
    header { padding: 2rem 1.5rem 1.25rem; text-align: center; border-bottom: 1px solid var(--border); background: #181715; }
    header h1 { font-size: 1.9rem; font-weight: normal; letter-spacing: 0.08em; color: var(--accent); margin-bottom: 0.4rem; }
    header p { font-family: system-ui, sans-serif; color: var(--muted); font-size: 0.95rem; }

    /* Navigation Bar */
    .room-nav { display: flex; flex-wrap: wrap; justify-content: center; align-items: center; gap: 0.6rem; padding: 1.2rem; border-bottom: 1px solid var(--border); background: #181715; }
    .room-btn { font-family: system-ui, sans-serif; background: var(--card-bg); color: var(--text); border: 1px solid var(--border); padding: 0.6rem 1.25rem; border-radius: 999px; cursor: pointer; font-size: 0.9rem; transition: all 0.2s; }
    .room-btn.active, .room-btn:hover { background: var(--accent); color: #141413; border-color: var(--accent); font-weight: 600; }

    .nav-divider { width: 1px; height: 24px; background: var(--border); margin: 0 0.35rem; }
    .quiz-nav-btn { font-family: system-ui, sans-serif; background: transparent; color: var(--accent); border: 1px solid var(--accent); padding: 0.6rem 1.35rem; border-radius: 999px; cursor: pointer; font-size: 0.9rem; font-weight: 600; transition: all 0.2s; }
    .quiz-nav-btn.active, .quiz-nav-btn:hover { background: var(--accent); color: #141413; }

    main { max-width: 1200px; margin: 0 auto; padding: 2rem 1.5rem; }
    .room-header { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 1.5rem; border-bottom: 1px solid var(--border); padding-bottom: 0.75rem; }
    .room-count { font-family: system-ui, sans-serif; font-size: 0.85rem; color: var(--muted); }

    /* Artwork Grid */
    .art-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 1.5rem; }
    .art-card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 8px; overflow: hidden; cursor: pointer; transition: transform 0.2s, border-color 0.2s, box-shadow 0.2s; display: flex; flex-direction: column; }
    .art-card:hover { transform: translateY(-4px); border-color: var(--accent); box-shadow: 0 10px 24px rgba(0,0,0,0.45); }
    .art-thumb-wrap { width: 100%; aspect-ratio: 4 / 3; background: #0e0e0d; overflow: hidden; display: flex; align-items: center; justify-content: center; }
    .art-thumb { width: 100%; height: 100%; object-fit: cover; display: block; transition: transform 0.3s ease; }
    .art-card:hover .art-thumb { transform: scale(1.04); }
    .art-card-caption { padding: 0.9rem 1rem; text-align: center; font-size: 0.95rem; line-height: 1.35; }

    /* Artwork Modal */
    .modal-backdrop { position: fixed; inset: 0; background: rgba(0,0,0,0.85); display: none; align-items: center; justify-content: center; padding: 1.5rem; z-index: 1000; backdrop-filter: blur(4px); }
    .modal-backdrop.open { display: flex; }
    .modal-content { background: var(--card-bg); border: 1px solid var(--accent); border-radius: 10px; max-width: 960px; width: 100%; max-height: 90vh; overflow-y: auto; display: grid; grid-template-columns: 1.1fr 1fr; position: relative; box-shadow: 0 20px 50px rgba(0,0,0,0.7); }
    @media (max-width: 760px) { .modal-content { grid-template-columns: 1fr; } }
    .modal-image-col { background: #0b0b0a; display: flex; align-items: center; justify-content: center; padding: 1.5rem; min-height: 320px; }
    .modal-image-col img { max-width: 100%; max-height: 75vh; object-fit: contain; border-radius: 4px; }
    .modal-details-col { padding: 2.25rem 2rem; display: flex; flex-direction: column; justify-content: flex-start; overflow-y: auto; }
    .modal-room-tag { font-family: system-ui, sans-serif; text-transform: uppercase; letter-spacing: 0.12em; font-size: 0.75rem; color: var(--accent); margin-bottom: 0.5rem; }
    .modal-title { font-size: 1.6rem; font-weight: normal; margin-bottom: 1rem; line-height: 1.25; }
    .modal-meta { font-family: system-ui, sans-serif; border-top: 1px solid var(--border); border-bottom: 1px solid var(--border); padding: 0.85rem 0; margin-bottom: 1.25rem; font-size: 0.92rem; line-height: 1.7; }
    .modal-meta span.label { color: var(--muted); display: inline-block; width: 110px; }
    .modal-description { line-height: 1.65; font-size: 0.98rem; color: #e4dfd7; white-space: pre-line; }
    .close-btn { position: absolute; top: 0.85rem; right: 1rem; background: #141413; color: var(--text); border: 1px solid var(--border); width: 36px; height: 36px; border-radius: 50%; font-size: 1.25rem; cursor: pointer; z-index: 2; }
    .close-btn:hover { border-color: var(--accent); color: var(--accent); }

    /* Trivia Quiz Styles */
    .quiz-container { max-width: 760px; margin: 0 auto; }
    .quiz-score-banner { background: var(--card-bg); border: 1px solid var(--accent); border-radius: 10px; padding: 1.5rem; text-align: center; margin-bottom: 2rem; }
    .quiz-score-banner h3 { font-size: 1.65rem; color: var(--accent); font-weight: normal; margin-bottom: 0.35rem; }
    .quiz-score-banner p { font-family: system-ui, sans-serif; color: var(--muted); font-size: 0.95rem; }

    .quiz-card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 10px; padding: 1.5rem 1.75rem; margin-bottom: 1.25rem; transition: border-color 0.2s; }
    .quiz-card.is-correct { border-color: var(--success); }
    .quiz-card.is-wrong { border-color: var(--danger); }

    .quiz-q-num { font-family: system-ui, sans-serif; font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.1em; color: var(--accent); margin-bottom: 0.4rem; }
    .quiz-q-text { font-size: 1.15rem; line-height: 1.45; margin-bottom: 1.1rem; }

    .quiz-options { display: grid; gap: 0.65rem; }
    .quiz-option { font-family: system-ui, sans-serif; display: flex; align-items: center; gap: 0.75rem; padding: 0.75rem 1rem; background: #171614; border: 1px solid var(--border); border-radius: 8px; cursor: pointer; font-size: 0.95rem; transition: all 0.15s; }
    .quiz-option:hover { border-color: var(--accent); }
    .quiz-option input[type="radio"] { accent-color: var(--accent); width: 16px; height: 16px; cursor: pointer; }
    .quiz-option.opt-correct { border-color: var(--success); background: var(--success-bg); color: #dcfce7; }
    .quiz-option.opt-wrong { border-color: var(--danger); background: var(--danger-bg); color: #fee2e2; }

    .quiz-feedback { font-family: system-ui, sans-serif; font-size: 0.88rem; margin-top: 0.85rem; }
    .quiz-feedback.ok { color: var(--success); }
    .quiz-feedback.bad { color: var(--danger); }

    .quiz-actions { display: flex; justify-content: center; gap: 1rem; margin-top: 2rem; }
    .primary-btn { font-family: system-ui, sans-serif; background: var(--accent); color: #141413; border: none; padding: 0.85rem 2.25rem; border-radius: 999px; font-size: 1rem; font-weight: 600; cursor: pointer; transition: opacity 0.2s; }
    .primary-btn:hover { opacity: 0.9; }
    .secondary-btn { font-family: system-ui, sans-serif; background: var(--card-bg); color: var(--text); border: 1px solid var(--border); padding: 0.85rem 1.75rem; border-radius: 999px; font-size: 0.95rem; cursor: pointer; }
    .secondary-btn:hover { border-color: var(--accent); color: var(--accent); }
  </style>
</head>
<body>
  <header>
    <h1>321 YORK ART & SCULPTURE COLLECTION</h1>
    <p>Select a room below, then click any artwork to view its title, artist, date, and description.</p>
  </header>
  <nav id="roomNav" class="room-nav"></nav>
  <main id="mainContent"></main>

  <div id="artModal" class="modal-backdrop" onclick="if(event.target.id==='artModal')closeModal()">
    <div class="modal-content">
      <button class="close-btn" onclick="closeModal()">&times;</button>
      <div class="modal-image-col"><img id="modalImg" src="" alt="" referrerpolicy="no-referrer"></div>
      <div class="modal-details-col">
        <div id="modalRoom" class="modal-room-tag"></div>
        <h3 id="modalTitle" class="modal-title"></h3>
        <div class="modal-meta">
          <div><span class="label">Artist:</span> <strong id="modalArtist"></strong></div>
          <div><span class="label">Date:</span> <span id="modalDate"></span></div>
        </div>
        <div id="modalDesc" class="modal-description"></div>
      </div>
    </div>
  </div>

  <script>
    const rawB64 = "${safeBase64Json}";
    const jsonBytes = Uint8Array.from(atob(rawB64), c => c.charCodeAt(0));
    const data = JSON.parse(new TextDecoder().decode(jsonBytes));
    const resolvedImages = {};
    let shuffledQuiz = [];

    function init() {
      const nav = document.getElementById('roomNav');
      data.rooms.forEach(room => {
        const btn = document.createElement('button');
        btn.className = 'room-btn';
        btn.textContent = room;
        btn.onclick = () => selectRoom(room);
        nav.appendChild(btn);
      });

      // Add Trivia Quiz button to the navigation bar
      const divider = document.createElement('div');
      divider.className = 'nav-divider';
      nav.appendChild(divider);

      const quizBtn = document.createElement('button');
      quizBtn.className = 'quiz-nav-btn';
      quizBtn.id = 'quizNavBtn';
      quizBtn.textContent = 'Trivia Quiz';
      quizBtn.onclick = () => openTriviaQuiz();
      nav.appendChild(quizBtn);

      if (data.rooms.length > 0) selectRoom(data.rooms[0]);
    }

    function selectRoom(roomName) {
      document.querySelectorAll('.room-btn').forEach(b => {
        b.classList.toggle('active', b.textContent === roomName);
      });
      const quizBtn = document.getElementById('quizNavBtn');
      if (quizBtn) quizBtn.classList.remove('active');

      const list = data.artworks.filter(a => a.room === roomName);
      const main = document.getElementById('mainContent');
      main.innerHTML =
        '<div class="room-header"><h2>' + escapeHtml(roomName) + '</h2>' +
        '<span class="room-count">' + list.length + (list.length === 1 ? ' Artwork' : ' Artworks') + '</span></div>' +
        '<div class="art-grid" id="grid"></div>';
      const grid = document.getElementById('grid');

      list.forEach(art => {
        const card = document.createElement('div');
        card.className = 'art-card';
        card.onclick = () => openModal(art);

        const wrap = document.createElement('div');
        wrap.className = 'art-thumb-wrap';

        const img = document.createElement('img');
        img.className = 'art-thumb';
        img.id = 'img-' + art.index;
        img.alt = art.title;
        img.referrerPolicy = 'no-referrer';

        const cachedSrc = resolvedImages[art.index];
        if (cachedSrc) {
          img.src = cachedSrc;
        } else if (art.imageUrl) {
          img.src = art.imageUrl;
          img.onerror = () => handleImageFallback(art, img, 1);
        }

        const cap = document.createElement('div');
        cap.className = 'art-card-caption';
        cap.textContent = art.title;

        wrap.appendChild(img);
        card.appendChild(wrap);
        card.appendChild(cap);
        grid.appendChild(card);
      });
    }

    /**
     * Fisher-Yates shuffle to randomize the 4 answer choices per question
     */
    function shuffleArray(arr) {
      const copy = arr.slice();
      for (let i = copy.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const temp = copy[i];
        copy[i] = copy[j];
        copy[j] = temp;
      }
      return copy;
    }

    function openTriviaQuiz() {
      document.querySelectorAll('.room-btn').forEach(b => b.classList.remove('active'));
      const quizBtn = document.getElementById('quizNavBtn');
      if (quizBtn) quizBtn.classList.add('active');

      const main = document.getElementById('mainContent');
      const questions = data.trivia || [];

      if (questions.length === 0) {
        main.innerHTML =
          '<div class="quiz-container"><div class="quiz-score-banner">' +
          '<h3>No Trivia Questions Found</h3>' +
          '<p>Add questions to the "Trivia" tab in your spreadsheet to start the quiz.</p>' +
          '</div></div>';
        return;
      }

      // Randomize answer choices for every question
      shuffledQuiz = questions.map((q, idx) => ({
        index: idx,
        question: q.question,
        correctAnswer: q.correctAnswer,
        choices: shuffleArray(q.choices)
      }));

      let html =
        '<div class="room-header"><h2>House & Collection Trivia</h2>' +
        '<span class="room-count">' + shuffledQuiz.length + (shuffledQuiz.length === 1 ? ' Question' : ' Questions') + '</span></div>' +
        '<div class="quiz-container">' +
        '<div id="quizScoreBanner"></div>' +
        '<form id="quizForm" onsubmit="submitTriviaQuiz(event)">';

      shuffledQuiz.forEach((item, qIdx) => {
        html +=
          '<div class="quiz-card" id="qcard-' + qIdx + '">' +
          '<div class="quiz-q-num">Question ' + (qIdx + 1) + ' of ' + shuffledQuiz.length + '</div>' +
          '<div class="quiz-q-text">' + escapeHtml(item.question) + '</div>' +
          '<div class="quiz-options">';

        item.choices.forEach((choice, cIdx) => {
          const inputId = 'q' + qIdx + '_opt' + cIdx;
          html +=
            '<label class="quiz-option" id="lbl-' + inputId + '" for="' + inputId + '">' +
            '<input type="radio" id="' + inputId + '" name="question_' + qIdx + '" value="' + escapeHtml(choice) + '">' +
            '<span>' + escapeHtml(choice) + '</span>' +
            '</label>';
        });

        html +=
          '</div>' +
          '<div class="quiz-feedback" id="qfb-' + qIdx + '"></div>' +
          '</div>';
      });

      html +=
        '<div class="quiz-actions">' +
        '<button type="submit" class="primary-btn" id="submitQuizBtn">Submit Answers</button>' +
        '<button type="button" class="secondary-btn" onclick="openTriviaQuiz()">Shuffle & Restart</button>' +
        '</div>' +
        '</form></div>';

      main.innerHTML = html;
    }

    function submitTriviaQuiz(e) {
      e.preventDefault();
      let score = 0;
      const total = shuffledQuiz.length;

      shuffledQuiz.forEach((item, qIdx) => {
        const selected = document.querySelector('input[name="question_' + qIdx + '"]:checked');
        const userChoice = selected ? selected.value : '';
        const card = document.getElementById('qcard-' + qIdx);
        const feedback = document.getElementById('qfb-' + qIdx);

        card.classList.remove('is-correct', 'is-wrong');

        // Highlight the choices and lock inputs
        item.choices.forEach((choice, cIdx) => {
          const inputId = 'q' + qIdx + '_opt' + cIdx;
          const lbl = document.getElementById('lbl-' + inputId);
          const inp = document.getElementById(inputId);
          lbl.classList.remove('opt-correct', 'opt-wrong');
          if (inp) inp.disabled = true;

          if (choice === item.correctAnswer) {
            lbl.classList.add('opt-correct');
          } else if (userChoice === choice) {
            lbl.classList.add('opt-wrong');
          }
        });

        if (userChoice === item.correctAnswer) {
          score++;
          card.classList.add('is-correct');
          feedback.className = 'quiz-feedback ok';
          feedback.textContent = 'Correct!';
        } else {
          card.classList.add('is-wrong');
          feedback.className = 'quiz-feedback bad';
          feedback.textContent = userChoice
            ? 'Incorrect — Correct answer: ' + item.correctAnswer
            : 'Unanswered — Correct answer: ' + item.correctAnswer;
        }
      });

      const pct = Math.round((score / total) * 100);
      const banner = document.getElementById('quizScoreBanner');
      banner.innerHTML =
        '<div class="quiz-score-banner">' +
        '<h3>You scored ' + score + ' out of ' + total + ' (' + pct + '%)</h3>' +
        '<p>Review your answers below or click "Shuffle & Restart" to try again.</p>' +
        '</div>';

      const submitBtn = document.getElementById('submitQuizBtn');
      if (submitBtn) submitBtn.style.display = 'none';

      banner.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    function handleImageFallback(art, imgEl, attempt) {
      if (attempt === 1 && art.altDriveUrl) {
        imgEl.onerror = () => handleImageFallback(art, imgEl, 2);
        imgEl.src = art.altDriveUrl;
        return;
      }

      if (art.driveFileId && typeof google !== 'undefined' && google.script && google.script.run) {
        imgEl.onerror = null;
        google.script.run
          .withSuccessHandler(function(dataUri) {
            if (dataUri) {
              resolvedImages[art.index] = dataUri;
              imgEl.src = dataUri;
              const modalImg = document.getElementById('modalImg');
              if (modalImg && modalImg.dataset.activeIndex == art.index) {
                modalImg.src = dataUri;
              }
            }
          })
          .fetchDriveImageBase64(art.driveFileId);
      }
    }

    function openModal(art) {
      const modalImg = document.getElementById('modalImg');
      const currentThumb = document.getElementById('img-' + art.index);
      modalImg.dataset.activeIndex = art.index;
      modalImg.src = resolvedImages[art.index] || (currentThumb ? currentThumb.src : art.imageUrl) || '';
      modalImg.alt = art.title;

      document.getElementById('modalRoom').textContent = art.room;
      document.getElementById('modalTitle').textContent = art.title;
      document.getElementById('modalArtist').textContent = art.artist;
      document.getElementById('modalDate').textContent = art.datePainted;
      document.getElementById('modalDesc').textContent = art.description;
      document.getElementById('artModal').classList.add('open');
    }

    function closeModal() {
      document.getElementById('artModal').classList.remove('open');
    }

    function escapeHtml(str) {
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') closeModal();
    });

    init();
  </script>
</body>
</html>`;
}
