/* Mail Link – henter direktelenke (deeplink) til valgt epost */

let current = { link: "", subject: "" };

Office.onReady(() => {
  document.getElementById("copy").onclick = () => copyText(current.link, "Lenke kopiert");
  document.getElementById("copyMd").onclick = () =>
    copyText(`[${escapeMd(current.subject)}](${current.link})`, "Markdown-lenke kopiert");
  document.getElementById("copyHtml").onclick = copyRichLink;
  document.getElementById("open").onclick = () => current.link && window.open(current.link, "_blank");

  refresh();
  // Når panelet er festet (pinnet) oppdateres det automatisk når du bytter epost.
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, refresh);
  } catch (e) { /* eldre klienter uten støtte */ }
});

function refresh() {
  const item = Office.context.mailbox.item;
  if (!item || !item.itemId) {
    setView("", "", "Velg en epost for å hente lenken.");
    return;
  }
  try {
    current = { link: buildLink(item.itemId), subject: item.subject || "(uten emne)" };
    setView(current.link, current.subject, "");
  } catch (e) {
    setView("", "", "Kunne ikke lage lenke: " + e.message, true);
  }
}

function buildLink(itemId) {
  const mb = Office.context.mailbox;
  const restId = mb.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
  const host = mb.userProfile.accountType === "outlookCom"
    ? "https://outlook.live.com"
    : "https://outlook.office.com";
  return `${host}/owa/?ItemID=${encodeURIComponent(restId)}&exvsurl=1&viewmodel=ReadMessageItem`;
}

function setView(link, subject, message, isError) {
  document.getElementById("subject").textContent = subject || "";
  document.getElementById("link").value = link;
  if (!link) current = { link: "", subject: "" };
  status(message, isError);
}

function status(msg, isError) {
  const el = document.getElementById("status");
  el.textContent = msg || "";
  el.className = isError ? "err" : "";
}

function escapeMd(s) { return s.replace(/([\[\]\\])/g, "\\$1"); }
function escapeHtml(s) {
  return s.replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

async function copyText(text, okMsg) {
  if (!text) return status("Ingen lenke å kopiere.", true);
  try {
    await navigator.clipboard.writeText(text);
    status(okMsg);
  } catch (e) {
    // Fallback hvis clipboard-API er blokkert i panelet
    const ta = document.getElementById("link");
    const old = ta.value;
    ta.value = text; ta.select();
    const ok = document.execCommand("copy");
    ta.value = old;
    status(ok ? okMsg : "Kunne ikke kopiere – merk teksten og trykk Ctrl+C.", !ok);
  }
}

async function copyRichLink() {
  if (!current.link) return status("Ingen lenke å kopiere.", true);
  const html = `<a href="${current.link}">${escapeHtml(current.subject)}</a>`;
  try {
    await navigator.clipboard.write([new ClipboardItem({
      "text/html": new Blob([html], { type: "text/html" }),
      "text/plain": new Blob([current.link], { type: "text/plain" })
    })]);
    status("Lenke med tittel kopiert");
  } catch (e) {
    copyText(current.link, "Rik kopiering ikke støttet – vanlig lenke kopiert");
  }
}
