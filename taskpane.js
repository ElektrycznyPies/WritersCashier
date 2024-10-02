/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("calculate").onclick = calculateEarnings;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // akapit na końcu dokumentu
    const paragraph = context.document.body.insertParagraph("Made by ElektrycznyPies", Word.InsertLocation.end);

    // kolor czcionki akapitu na niebieski
    paragraph.font.color = "blue";

    await context.sync();
  });
}

async function calculateEarnings() {
  return Word.run(async (context) => {
    // Pobranie stawki i liczby znaków, za które należy się stawka
    const rate = parseFloat(document.getElementById("rate").value);
    const perChars = parseFloat(document.getElementById("perChars").value);

    // Pobranie liczby znaków w dokumencie
    const body = context.document.body;
    body.load("text");

    await context.sync();

    const numCharacters = body.text.length;

    // Obliczenie zarobków
    const earnings = (numCharacters / perChars) * rate;

    // Wyświetlenie liczby znaków i zarobków
    document.getElementById("outputCharsNumber").textContent = `Number of characters in the document: ${numCharacters}`;
    document.getElementById("outputCashEarned").textContent = `MONEY EARNED: ${earnings.toFixed(2)}`;

    await context.sync();
  });
}
