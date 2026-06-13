/**
 * ============================================================
 * Speaker card from template generation with Google Sheets
 * and Google Slides
 * ============================================================
 *  SETUP:
 *  1. Speaker Details — this script is container-bound, so create it
 *     via Extensions → Apps Script from your speakers Sheet.
 *     Column names must be in Row 1. Every column
 *     becomes a placeholder automatically: a column headed
 *     `speaker_name` is written into the template wherever you
 *     put {{speaker_name}}. No column is special — except the
 *     photo one (see CONFIG.PHOTO_COLUMN below), which holds a
 *     URL Slides can fetch rather than text. For Drive-hosted
 *     images, share them ("anyone with link") and use:
 *       https://drive.google.com/uc?export=view&id=FILE_ID
 *
 *  2. THE TEMPLATE — a Google Slides deck where SLIDE 1 is your
 *     card design. For each column you want shown, use a text
 *     box containing {{thatColumnName}}. Matching is case-
 *     insensitive ({{name}} hits a "Name" column), but otherwise
 *     the placeholder must match the column text as written, so a
 *     column "Speaker Name" needs {{Speaker Name}}.
 *     For the photo, insert ANY placeholder image and set its
 *     alt text (Format options → Alt text) to: speaker-photo
 *     Tip: set the deck's page size to your card size first
 *     (File → Page setup → Custom → 1200×630 px) so exports
 *     come out at the right aspect ratio.
 *
 *  3. CONFIG — paste your IDs below.
 *
 *  4. SCOPES — first run will ask for permissions. The export
 *     function calls the Slides REST API directly, so if you
 *     get a 403, enable the "Google Slides API" under
 *     Services (+) in the Apps Script editor.
 *
 *  Then: reload the Sheet and use the 🦥 menu. That's it.
 */

const CONFIG = {
  // The Slides deck whose first slide is your card template 
  // (get ID from URL of the Google Slides - and make sure you own this file and have edit)
  TEMPLATE_PRESENTATION_ID: 'PASTE_SLIDES_ID_HERE',

  // Drive folder where the exported PNGs should land
  OUTPUT_FOLDER_ID: 'PASTE_FOLDER_ID_HERE',

  // The tab in this spreadsheet holding speaker data
  SHEET_NAME: 'Speakers',

  // Which column holds the photo URL.
  // Must be a publicly accessible URL that goes straight to the image
  PHOTO_COLUMN: 'photo',

  // Alt text that marks the photo placeholder in the template
  PHOTO_ALT_TEXT: 'speaker-photo',
};

/** Adds the menu. Reload the Sheet after first save. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🦥 Lazy Organiser')
    .addItem('1. Generate speaker cards', 'createSpeakerCards')
    .addItem('2. Export cards as PNGs', 'exportCardsToDrive')
    .addSeparator()
    .addItem('Reset deck (delete generated slides)', 'resetDeck')
    .addToUi();
}

/**
 * STEP 1 — One slide per speaker.
 * Duplicates the template slide, swaps placeholders, swaps photo.
 */
function createSpeakerCards() {
  const speakers = readSpeakers_();
  const deck = SlidesApp.openById(CONFIG.TEMPLATE_PRESENTATION_ID);
  const template = deck.getSlides()[0];

  speakers.forEach((speaker) => {
    const slide = template.duplicate();
    // duplicate() drops the copy right after the template, which
    // reverses the order — move each one to the end instead so
    // slide order matches sheet order (the export step relies on this).
    slide.move(deck.getSlides().length);

    // Walk every column. The photo column gets swapped as an image;
    // everything else is a text placeholder: header "foo" ➜ {{foo}}.
    Object.keys(speaker).forEach((column) => {
      const value = speaker[column];

      if (normalise_(column) === normalise_(CONFIG.PHOTO_COLUMN)) {
        // Photo placeholder — found by alt text, replaced in place
        // (keeps the size/position/crop you set in the template).
        if (value) {
          const photo = slide
            .getImages()
            .find((img) => img.getDescription() === CONFIG.PHOTO_ALT_TEXT);
          if (photo) photo.replace(value);
        }
      } else {
        // matchCase=false ➜ {{name}} in the template still matches a
        // "Name" column. The text otherwise matches the header as written.
        slide.replaceAllText(
          `{{${column}}}`,
          value == null ? '' : String(value),
          false
        );
      }
    });
  });

  deck.saveAndClose();
  toast_(`Generated ${speakers.length} cards. Step 2 when ready.`);
}

/**
 * STEP 2 — Export every generated slide as a PNG.
 * Uses the Slides API thumbnail endpoint: up to 1600px wide,
 * no PDF-splitting nonsense. NOTE: the contentUrl it returns
 * expires after ~30 minutes, so we fetch the blob immediately.
 */
function exportCardsToDrive() {
  const speakers = readSpeakers_();
  const deck = SlidesApp.openById(CONFIG.TEMPLATE_PRESENTATION_ID);
  const slides = deck.getSlides();
  const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const token = ScriptApp.getOAuthToken();

  // slides[0] is the template; generated cards start at index 1
  // and match the sheet row order (see the move() in step 1).
  speakers.forEach((speaker, i) => {
    const slide = slides[i + 1];
    if (!slide) return;

    const url =
      `https://slides.googleapis.com/v1/presentations/` +
      `${CONFIG.TEMPLATE_PRESENTATION_ID}/pages/${slide.getObjectId()}` +
      `/thumbnail?thumbnailProperties.thumbnailSize=LARGE`;

    const meta = JSON.parse(
      UrlFetchApp.fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      }).getContentText()
    );

    const blob = UrlFetchApp.fetch(meta.contentUrl)
      .getBlob()
      .setName(`card-${slugify_(pickLabel_(speaker) || `speaker-${i + 1}`)}.png`);

    folder.createFile(blob);
  });

  toast_(`Exported ${speakers.length} PNGs to Drive. Go home.`);
}

/**
 * HOUSEKEEPING — deletes everything except the template slide,
 * so you can regenerate without a deck full of stale cards.
 */
function resetDeck() {
  const deck = SlidesApp.openById(CONFIG.TEMPLATE_PRESENTATION_ID);
  deck.getSlides().slice(1).forEach((slide) => slide.remove());
  deck.saveAndClose();
  toast_('Deck reset. Only the template remains.');
}

// ---------- helpers ----------

/** Loose comparison key — lowercased, with spaces/underscores stripped. */
function normalise_(text) {
  return String(text).toLowerCase().replace(/[\s_]+/g, '');
}

/**
 * Picks a sensible filename basis: the first non-photo column's
 * value (usually the speaker's name, whatever it's called).
 */
function pickLabel_(speaker) {
  const labelColumn = Object.keys(speaker).find(
    (column) => normalise_(column) !== normalise_(CONFIG.PHOTO_COLUMN)
  );
  return labelColumn ? speaker[labelColumn] : '';
}

/**
 * Reads the sheet into an array of objects keyed by header row.
 * Headers keep their original text (just trimmed) so a column
 * called "Speaker Name" maps to the placeholder {{Speaker Name}}.
 */
function readSpeakers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.SHEET_NAME
  );
  if (!sheet) throw new Error(`No sheet named "${CONFIG.SHEET_NAME}".`);

  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows
    .filter((row) => row.some((cell) => cell !== ''))
    .map((row) =>
      Object.fromEntries(
        headers.map((h, i) => [String(h).trim(), row[i]])
      )
    );
}

/** "Ada Lovelace" → "ada-lovelace" for tidy filenames. */
function slugify_(text) {
  return String(text)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '');
}

/** Non-blocking notification in the Sheet UI. */
function toast_(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '🦥');
}
