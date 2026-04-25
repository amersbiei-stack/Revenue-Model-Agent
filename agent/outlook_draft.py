"""Outlook COM helpers — build HTML mail with CID-embedded PNGs, save to Drafts."""
import os
import time
from pathlib import Path
from tempfile import NamedTemporaryFile

import win32com.client as win32

from agent.logging_setup import get_logger

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
OL_MAIL_ITEM = 0
OL_FORMAT_HTML = 2


def _get_outlook_with_retry(log):
    """Dispatch + CreateItem with retry. Outlook auto-launches the first
    time we Dispatch, but for a few seconds after launch it can reject
    COM calls with E_ABORT (0x80004004 'Operation aborted'). Sleep and
    retry rather than failing the whole step.
    """
    last_error = None
    for attempt in range(1, 6):
        try:
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(OL_MAIL_ITEM)
            return outlook, mail
        except Exception as e:
            last_error = e
            log.warning(
                "Outlook CreateItem failed (attempt %d/5): %s — sleeping before retry",
                attempt, e,
            )
            time.sleep(3 * attempt)  # 3s, 6s, 9s, 12s, 15s
    raise RuntimeError(
        "Could not create an Outlook draft after 5 retries. "
        "Make sure Outlook desktop is installed, signed in, and not "
        "blocked by a security prompt. Original error: " + str(last_error)
    ) from last_error


def create_draft(
    to: str,
    subject: str,
    html_body: str,
    cid_pngs: dict | None = None,
) -> str:
    """Create an Outlook mail item, attach PNGs as CID inline images, save to Drafts.

    cid_pngs maps CID strings (e.g. 'mtd_chart') to raw PNG bytes. The HTML body
    must reference them via <img src='cid:mtd_chart'>.

    Returns the draft's EntryID (useful for logging or later retrieval).
    Does NOT call .Send() — explicit reviewer step.
    """
    log = get_logger()
    outlook, mail = _get_outlook_with_retry(log)
    mail.To = to
    mail.Subject = subject
    mail.BodyFormat = OL_FORMAT_HTML

    temp_paths: list[str] = []
    try:
        if cid_pngs:
            for cid, png_bytes in cid_pngs.items():
                tmp = NamedTemporaryFile(delete=False, suffix=".png")
                tmp.write(png_bytes)
                tmp.close()
                temp_paths.append(tmp.name)
                att = mail.Attachments.Add(tmp.name)
                att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
                log.debug("Attached CID image: %s (%d bytes)", cid, len(png_bytes))

        mail.HTMLBody = html_body
        mail.Save()
        entry_id = mail.EntryID
        log.info("Outlook draft saved (EntryID=%s)", entry_id)
        return entry_id
    finally:
        for p in temp_paths:
            try:
                os.unlink(p)
            except OSError:
                pass

